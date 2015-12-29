#ifndef EXCEL_H
#define EXCEL_H

#include <QObject>

#include <QWidget>
#include <QFileDialog>
#include <QScopedPointer>
#include <QFile>
#include <QtConcurrentRun>
#include <QFutureWatcher>
#include <QFuture>
#include <QProgressDialog>
#include <QAxObject>
#include <QMessageBox>
#include <QStyledItemDelegate>
#include <QPainter>
#include <shlobj.h> //для использования QAxObject в отдельном потоке

#include "defines.h"
#include "address.h"
#include "xlsparser.h"
#include "database.h"


const QString FileName="Log1.csv";

class Excel : public QObject
{
    Q_OBJECT
public:
    explicit Excel(QObject *parent = 0);
    ~Excel();

public slots:
    void runThreadOpen(QString openFilename);
    void runThreadOpenCsv(QString openFilename);
    void search();
    void closeLog();

signals:
    void headReaded(QString sheet, MapAddressElementPosition head);
    void rowReaded(QString sheet, int nRow, QStringList row);
    void countRows(QString sheet, int count);
    void sheetsReaded(QStringList sheets);

    void working(); //начало чтения файла
    void finished(); //чтение файла окончено
    void searching();

    void toDebug(QString objName, QString mes);
    void toFile(QString objName, QString mes);

    void findRowInBase(QString sheetName, int nRow, Address addr);

private slots:
    void onRowRead(const QString &sheet, const int &nRow, QStringList &row);
    void onHeadRead(const QString &sheet, QStringList &head);

    //parser signal-slots
    void onRowParsed(QString sheet, int nRow, Address a);

    void onProcessOfOpenFinished();//после того окончили с открытием excel документа

    void onDebug(QString objName, QString mes);
    void onFile(QString objName, QString mes);

    void onFounedAddress(QString sheetName, int nRow, Address addr);
    void onNotFounedAddress(QString sheetName, int nRow, Address addr);


private:
    XlsParser                   *_parser;
    QHash<QString, int>          _sheetIndex;
    QFutureWatcher<QVariant>     _futureWatcher;
    QFutureWatcher<ListAddress>     _futureWatcherS;
    QMap<QString, MapAddressElementPosition> _mapHead;
    QMap<QString, MapAddressElementPosition> _mapPHead;
    QMap<QString, ListAddress > _data2;
    ExcelDocument _data;
    QThread *_thread;
    QFile _logFile;
    QTextStream _logStream;
    QVariant openExcelFile(QString filename, int maxCount);
    QVariant openCsvFile(QString filename, int maxCountRows);
};

#endif // EXCEL_H
