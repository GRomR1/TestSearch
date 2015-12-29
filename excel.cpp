#include "excel.h"

Excel::Excel(QObject *parent) : QObject(parent)
{
    _parser = new XlsParser;
    _thread = new QThread;
    _logFile.setFileName(FileName);
    _logFile.open(QIODevice::WriteOnly | QIODevice::Text | QIODevice::Truncate);
    _logStream.setDevice(&_logFile);

//    _parser->moveToThread(_thread);
    connect(this, SIGNAL(rowReaded(QString,int,QStringList)),
            _parser, SLOT(onReadRow(QString,int,QStringList)));
    connect(this, SIGNAL(headReaded(QString,MapAddressElementPosition)),
            _parser, SLOT(onReadHead(QString,MapAddressElementPosition)));
    connect(_parser, SIGNAL(rowParsed(QString,int,Address)),
            this, SLOT(onRowParsed(QString,int,Address)));
    connect(_thread, SIGNAL(finished()),
            _parser, SLOT(deleteLater()));
    connect(_thread, SIGNAL(finished()),
            _thread, SLOT(deleteLater()));
    _thread->start();

    connect(&_futureWatcher, SIGNAL(finished()),
            this, SLOT(onProcessOfOpenFinished()));

    connect(this, SIGNAL(toDebug(QString,QString)),
            this, SLOT(onDebug(QString,QString)));

    connect(this, SIGNAL(toFile(QString,QString)), SLOT(onFile(QString,QString)));

    emit toDebug("test","test");
}

Excel::~Excel()
{
    _thread->quit();
    _thread->wait();
    closeLog();
}

void Excel::closeLog()
{
    if(_logFile.isOpen())
        _logFile.close();
}

QVariant Excel::openCsvFile(QString filename, int maxCountRows)
{
    QString currTime=QDateTime::currentDateTime().toString("dd.MM.yyyy hh:mm:ss.zzz");
    qDebug() << "openCsvFile BEGIN"
             << QThread::currentThread()->currentThreadId()
             << currTime;
    QStringList data;
    QFile file1(filename);
    if (!file1.open(QIODevice::ReadOnly))
    {
        qDebug().noquote() << "Ошибка открытия для чтения";
        return data;
    }
    QTextStream in(&file1);
    QTextCodec *defaultTextCodec = QTextCodec::codecForName("Windows-1251");
    if (defaultTextCodec)
      in.setCodec(defaultTextCodec);
    int nRow=0;
    while (!in.atEnd())
    {
        QString line = in.readLine();
        if(!line.isEmpty())
        {
            if(nRow==0)
            {
                qDebug().noquote() << "H:" << line;
            }
            data.append(line);
            nRow++;
        }
        //оставливаем обработку если получено нужное количество строк
        if(maxCountRows>0 && nRow >= maxCountRows)
            break;
    }
    file1.close();

    qDebug() << "openCsvFile END"
             << QThread::currentThread()->currentThreadId()
             << QDateTime::currentDateTime().toString("dd.MM.yyyy hh:mm:ss.zzz");
    return QVariant::fromValue(data);
}

QVariant Excel::openExcelFile(QString filename, int maxCountRows)
{
    QString currTime=QDateTime::currentDateTime().toString("dd.MM.yyyy hh:mm:ss.zzz");
    qDebug() << "openExcelFile BEGIN"
             << QThread::currentThread()->currentThreadId()
             << currTime;

    HRESULT h_result = CoInitializeEx(NULL, COINIT_MULTITHREADED);
    switch(h_result)
    {
    case S_OK:
        qDebug("TestConnector: The COM library was initialized successfully on this thread");
        break;
    case S_FALSE:
        qWarning("TestConnector: The COM library is already initialized on this thread");
        break;
    case RPC_E_CHANGED_MODE:
        qWarning() << "TestConnector: A previous call to CoInitializeEx specified the concurrency model for this thread as multithread apartment (MTA)."
                   << " This could also indicate that a change from neutral-threaded apartment to single-threaded apartment has occurred";
        break;
    }

    // получаем указатель на Excel
    QScopedPointer<QAxObject> excel(new QAxObject("Excel.Application"));
    if(excel.isNull())
    {
        QString error="Cannot get Excel.Application";
        return QVariant(error);
    }

    QScopedPointer<QAxObject> workbooks(excel->querySubObject("Workbooks"));
    if(workbooks.isNull())
    {
        QString error="Cannot query Workbooks";
        return QVariant(error);
    }

    // на директорию, откуда грузить книгу
    QScopedPointer<QAxObject> workbook(workbooks->querySubObject(
                                           "Open(const QString&)",
                                           filename)
                                       );
    if(workbook.isNull())
    {
        QString error=
                QString("Cannot query workbook.Open(const %1)")
                .arg(filename);
        return QVariant(error);
    }

    QScopedPointer<QAxObject> sheets(workbook->querySubObject("Sheets"));
    if(sheets.isNull())
    {
        QString error="Cannot query Sheets";
        return QVariant(error);
    }

    int count = sheets->dynamicCall("Count()").toInt(); //получаем кол-во листов
    QStringList readedSheetNames;
    //читаем имена листов
    for (int i=1; i<=count; i++)
    {
        QScopedPointer<QAxObject> sheetItem(sheets->querySubObject("Item(int)", i));
        if(sheetItem.isNull())
        {
            QString error="Cannot query Item(int)"+QString::number(i);
            return QVariant(error);
        }
        readedSheetNames.append( sheetItem->dynamicCall("Name()").toString() );
        sheetItem->clear();
    }
    // проходим по всем листам документа
    int sheetNumber=0;
    ExcelDocument data;
    foreach (QString sheetName, readedSheetNames)
    {
        QScopedPointer<QAxObject> sheet(
                    sheets->querySubObject("Item(const QVariant&)",
                                           QVariant(sheetName))
                    );
        if(sheet.isNull())
        {
            QString error=
                    QString("Cannot query Item(const %1)")
                    .arg(sheetName);
            return QVariant(error);
        }

        QScopedPointer<QAxObject> usedRange(sheet->querySubObject("UsedRange"));
        QScopedPointer<QAxObject> usedRows(usedRange->querySubObject("Rows"));
        QScopedPointer<QAxObject> usedCols(usedRange->querySubObject("Columns"));
        int rows = usedRows->property("Count").toInt();
        int cols = usedCols->property("Count").toInt();

        //если на данном листе всего 1 строка (или меньше), то данный лист пуст => зачем он нам?
        if(rows>1)
        {
            //чтение данных
            ExcelSheet excelSheet; //таблица
            for(int row=1; row<=rows; row++)
            {
                QStringList rowList; //строка с данными
                for(int col=1; col<=cols; col++)
                {
                    QScopedPointer<QAxObject> cell (
                                sheet->querySubObject("Cells(QVariant,QVariant)",
                                                      row,
                                                      col)
                                );
                    QString result = cell->property("Value").toString();
                    rowList.append(result);
                    cell->clear();
                }
                excelSheet.append(rowList);

                //оставливаем обработку если получено нужное количество строк
                if(maxCountRows>0 && row-1 >= maxCountRows)
                    break;
            }
            data.insert(sheetName, excelSheet); //добавляем таблицу в документ
        }
        usedRange->clear();
        usedRows->clear();
        usedCols->clear();
        sheet->clear();
        sheetNumber++;
    }//end foreach _sheetNames

    sheets->clear();
    workbook->clear();
    workbooks->dynamicCall("Close()");
    workbooks->clear();
    excel->dynamicCall("Quit()");

    qDebug() << "openExcelFile END"
             << QThread::currentThread()->currentThreadId()
             << QDateTime::currentDateTime().toString("dd.MM.yyyy hh:mm:ss.zzz");
    return QVariant::fromValue(data);
}

void Excel::onProcessOfOpenFinished()
{
    QString currTime=QDateTime::currentDateTime().toString("dd.MM.yyyy hh:mm:ss.zzz");
    qDebug() << "ExcelWidget onProcessOfOpenFinished BEGIN"
             << currTime
             << this->thread()->currentThreadId();

    if(_futureWatcher.isFinished()
            && !_futureWatcher.isCanceled())
    {
        QVariant result = _futureWatcher.future().result();
        emit toDebug("excel", "finish");
        if(result.isValid())
        {
            emit toDebug("excel", "valid");
            if(result.canConvert< ExcelDocument >())
            {
                emit toDebug("excel", "canconvert ExcelDocument");
                ExcelDocument data = result.value<ExcelDocument>();
                foreach (QString sheetName, data.keys()) {
                    ExcelSheet rows=data[sheetName];
                    emit toDebug("excel", sheetName);
                    _data.insert(sheetName, rows);
                    int nRow=0;
                    foreach (QStringList row, rows) {
//                        emit toDebug("excel", row.join(";"));
                        if(nRow==0)
                            onHeadRead(sheetName, row);
                        else
                            onRowRead(sheetName, nRow-1, row);
                        nRow++;
                    }
                }
            }
            else if(result.canConvert< QString >())
            {
                QString error = result.toString();
                emit toDebug(objectName(), error);
            }
            else if(result.canConvert< QStringList >())
            {
                emit toDebug("excel", "canconvert QStringList");
                QStringList data = result.value<QStringList>();
                QString sheetName = "csv";
//                ExcelSheet sheet;
                int nRow=0;
                foreach (QString line, data) {
                    QStringList row = line.split(';');
                    if(nRow==0)
                        onHeadRead(sheetName, row);
                    else
                        onRowRead(sheetName, nRow-1, row);
//                    sheet.append(row);
                    nRow++;
                }
//                _data.insert(sheetName, sheet);
            }
        }
    }
    currTime=QDateTime::currentDateTime().toString("dd.MM.yyyy hh:mm:ss.zzz");
    qDebug() << "ExcelWidget onProcessOfOpenFinished END"
             << currTime
             << this->thread()->currentThreadId();
    emit finished();
}


void Excel::search()
{
    QString currTime=QDateTime::currentDateTime().toString("dd.MM.yyyy hh:mm:ss.zzz");
    qDebug() << "ExcelWidget search BEGIN"
             << currTime
             << this->thread()->currentThreadId();

    if(_data2.isEmpty())
    {
        return;
    }
    QList<QString> sheets = _data2.keys();
    foreach (QString sheet, sheets) {
        for(int i=0; i<_data2[sheet].size(); i++)
        {
            if(!_data2[sheet].at(i).isEmpty())
                emit findRowInBase(sheet, i, _data2[sheet].at(i));
            if(i%200==0)
                toDebug(objectName(), QString("Обработано строк %1").arg(QString::number(i)));
        }
    }

    emit searching();
    currTime=QDateTime::currentDateTime().toString("dd.MM.yyyy hh:mm:ss.zzz");
    qDebug() << "ExcelWidget search END"
             << currTime
             << this->thread()->currentThreadId();
}

void Excel::onFounedAddress(QString sheetName, int nRow, Address addr)
{
    _data2[sheetName][nRow]=addr;
//    emit toDebug("", QString::number(nRow+1)+";"+/*+QString::number(addr.getStreetId())
//                 +";"+QString::number(addr.getBuildId())+
//                 ";"+addr.toString(RAW)*/ addr.toCsv());
    emit toFile("", QString::number(nRow+1)+";"+addr.toCsv()+"\n");
}

void Excel::onNotFounedAddress(QString sheetName, int nRow, Address addr)
{
//    _data2[sheetName][nRow]=addr;
    Q_UNUSED(sheetName);
    Q_UNUSED(addr);
//    emit toDebug("", QString::number(nRow+1)+";!NOT_FOUND!");
    emit toFile("", QString::number(nRow+1)+";!NOT_FOUND!"+"\n");
}

void Excel::runThreadOpenCsv(QString openFilename)
{
    qDebug() << "ExcelWidget runThreadOpenCSV" << this->thread()->currentThreadId();

    QString name = QFileInfo(openFilename).fileName();

    QProgressDialog *dialog = new QProgressDialog;
    dialog->setWindowTitle(trUtf8("Открываю файл..."));
    dialog->setLabelText(trUtf8("Открывается файл \"%1\". Ожидайте ...")
                         .arg(name));
    dialog->setCancelButtonText(trUtf8("Отмена"));
    QObject::connect(dialog, SIGNAL(canceled()),
                     &_futureWatcher, SLOT(cancel()));
    QObject::connect(&_futureWatcher, SIGNAL(progressRangeChanged(int,int)),
                     dialog, SLOT(setRange(int,int)));
    QObject::connect(&_futureWatcher, SIGNAL(progressValueChanged(int)),
                     dialog, SLOT(setValue(int)));
    QObject::connect(&_futureWatcher, SIGNAL(finished()),
                     dialog, SLOT(deleteLater()));

    QFuture<QVariant> f1 = QtConcurrent::run(this,
                                             &Excel::openCsvFile,
                                             openFilename,
                                             0);
    // Start the computation.
    _futureWatcher.setFuture(f1);
    dialog->exec();

    emit working();
}

void Excel::runThreadOpen(QString openFilename)
{
    qDebug() << "ExcelWidget runThreadOpen" << this->thread()->currentThreadId();

    QString name = QFileInfo(openFilename).fileName();

    QProgressDialog *dialog = new QProgressDialog;
    dialog->setWindowTitle(trUtf8("Открываю файл..."));
    dialog->setLabelText(trUtf8("Открывается файл \"%1\". Ожидайте ...")
                         .arg(name));
    dialog->setCancelButtonText(trUtf8("Отмена"));
    QObject::connect(dialog, SIGNAL(canceled()),
                     &_futureWatcher, SLOT(cancel()));
    QObject::connect(&_futureWatcher, SIGNAL(progressRangeChanged(int,int)),
                     dialog, SLOT(setRange(int,int)));
    QObject::connect(&_futureWatcher, SIGNAL(progressValueChanged(int)),
                     dialog, SLOT(setValue(int)));
    QObject::connect(&_futureWatcher, SIGNAL(finished()),
                     dialog, SLOT(deleteLater()));

    QFuture<QVariant> f1 = QtConcurrent::run(this,
                                             &Excel::openExcelFile,
                                             openFilename,
                                             0);
    // Start the computation.
    _futureWatcher.setFuture(f1);
    dialog->exec();

    emit working();
}


void Excel::onRowRead(const QString &sheet, const int &nRow, QStringList &row)
{
//    emit toDebug(objectName(),
//                 QString("onRowRead \"%1\" \"%2\"")
//                 .arg(sheet)
//                 .arg(row.join(";")));
    _data[sheet].insert(nRow, row);
    emit rowReaded(sheet, nRow, row);
}

void Excel::onHeadRead(const QString &sheet, QStringList &head)
{
    emit toDebug(objectName(),
                 QString("Прочитана шапка у листа \"%1\"")
                 .arg(sheet));

    //если остуствует столбец с данными о улице (городе и пр.)
    QString colname = MapColumnNames[ STREET ];
    if(!head.contains(colname))
    {
        emit toDebug(objectName(),
                     QString("Ошибка. Отсутсвует столбец \"%1\"")
                     .arg(colname));
        return;
    }
    colname = MapColumnNames[ STREET_ID ];
    if(head.indexOf(colname)==-1)
    {
        head.append(colname);
    }
    colname = MapColumnNames[ BUILD_ID ];
    if(!head.contains(colname))
    {
        head.append(colname);
    }

    QMap<AddressElements, QString>::const_iterator it
            = MapColumnParsedNames.begin();
    while(it!=MapColumnParsedNames.end())
    {
        colname = it.value();
        head.append(colname);
        it++;
    }

    QList<AddressElements> keys = MapColumnParsedNames.keys();
    foreach (AddressElements ae, keys) {
        QString colname;
        int pos=0;
        colname = MapColumnNames.value(ae);
        pos = head.indexOf(colname);
        if(pos!=-1)
            _mapHead[sheet].insert(ae, pos);
        colname = MapColumnParsedNames.value(ae);
        pos = head.indexOf(colname);
        if(pos!=-1)
            _mapPHead[sheet].insert(ae, pos);
    }
    if(_mapHead.contains(sheet))
        emit headReaded(sheet, _mapHead[sheet]);
}

void Excel::onRowParsed(QString sheet, int nRow, Address a)
{
//    emit toDebug(objectName(),
//                 "ExcelWidget::onRowParsed "
//                 +sheet+" row:"
//                 +QString::number(nRow)+"\n"+a.toString(PARSED)
//                 );
    _data2[sheet].insert(nRow, a);
}

void Excel::onFile(QString objName, QString mes)
{
    Q_UNUSED(objName);
    _logStream << mes;
}

void Excel::onDebug(QString objName, QString mes)
{
    qDebug().noquote() << objName << mes;
}
