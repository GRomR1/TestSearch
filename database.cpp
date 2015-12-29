#include "database.h"

Database::Database(QObject *parent) :
    QObject(parent),
    _thread(nullptr),
    _model(nullptr),
    _connected(false)
{
    _thread = new QThread;
}

Database::~Database()
{
    if(_thread)
    {
        _thread->quit();
        _thread->wait();
        delete _thread;
    }
    if(_model!=nullptr)
        delete _model;
}

void Database::setBaseName(QString name)
{
    _baseName = name;
}

void Database::openBase()
{
    if(!_connected)
        createConnection();
    if(!_connected)
        return;
//    openTableToModel();
//    emit countRows(_model->rowCount());
    emit baseOpened();
}

ListAddress Database::search(QString sheetName, ListAddress addr)
{
    qDebug() << "Database::search" << this->thread()->currentThreadId()
             << sheetName << addr.size();
    for(int i=0; i<addr.size(); i++)
    {
        selectAddress(addr[i]);
    }
    return addr;
}

void Database::clear()
{
}

QSqlTableModel *Database::getModel()
{
    return _model;
}

void Database::openTableToModel()
{
    if(_model!=nullptr)
        delete _model;
    _model = new QSqlTableModel(this, _db);
    _model->setTable(TableName);
    _model->setEditStrategy(QSqlTableModel::OnManualSubmit);
    _model->select();

    if(_model->lastError().isValid())
        emit toDebug(objectName(),
                     "openModel:"+_model->lastError().text());
}

void Database::createConnection()
{
    if(_connected)
        return;

    _db = QSqlDatabase::addDatabase("QSQLITE");
    _db.setDatabaseName(_baseName);
    _db.setUserName("user");
    _db.setHostName("rt");
    _db.setPassword("user");
    QString str;
    if (!_db.open())
    {
        _connected=false;
        str+="Cannot open database:" + _db.lastError().text();
        emit toDebug(objectName(),
                     str);
    }
    else
    {
        _connected=true;
        str += "Success open base:" + _baseName;
        emit toDebug(objectName(),
                     str);
    }
}

void Database::createTable()
{
    QSqlQuery query;
    QString str =
            QString("CREATE TABLE IF NOT EXISTS %18 ( "
                    "'%1' INTEGER PRIMARY KEY NOT NULL, "
                    "'%2' TEXT, "
                    "'%3' TEXT, "
                    "'%4' TEXT, "
                    "'%5' TEXT, "
                    "'%6' TEXT, "
                    "'%7' TEXT, "
                    "'%8' TEXT, "
                    "'%9' TEXT, "
                    "'%10' TEXT, "
                    "'%11' TEXT, "
                    "'%12' TEXT, "
                    "'%13' TEXT, "
                    "'%14' TEXT, "
                    "'%15' TEXT, "
                    "'%16' TEXT, "
                    "'%17' TEXT "
                    ");")
            .arg(MapColumnNames[BUILD_ID])
            .arg(MapColumnNames[STREET])
            .arg(MapColumnNames[STREET_ID])
            .arg(MapColumnNames[KORP])
            .arg(MapColumnNames[BUILD])
            .arg(MapColumnNames[TYPE_OF_STREET])
            .arg(MapColumnNames[ADDITIONAL])
            .arg(MapColumnNames[TYPE_OF_CITY1])
            .arg(MapColumnNames[CITY1])
            .arg(MapColumnNames[TYPE_OF_CITY2])
            .arg(MapColumnNames[CITY2])
            .arg(MapColumnNames[DISTRICT])
            .arg(MapColumnNames[FSUBJ])
            .arg(MapColumnNames[RAW_ADDR])
            .arg(MapColumnNames[LITERA])
            .arg(MapColumnNames[CORRECT])
            .arg(MapColumnNames[TYPE_OF_FSUBJ])
            .arg(TableName);

    if(!query.exec(str))
        toDebug(objectName(),
                "Unable to create a table:\r\n"+query.lastError().text());
    else
        toDebug(objectName(),
                "Success create a table");

//    if(_model==nullptr)
//    {
//        _model = new QSqlTableModel;
//        _model->setTable("base");
//        _model->setEditStrategy(QSqlTableModel::OnManualSubmit);
//        _model->select();
//    }
}

void Database::insertAddress(Address a)
{
//    qDebug().noquote() << "Database insert" << QThread::currentThreadId();
//    QSqlQuery query(_db);
//    query.prepare("INSERT INTO base1 (BUILD_ID, BUILD, STREET_ID, STREET,"
//                  "TYPE_OF_STREET, KORP, LITERA, CORRECT, CITY1, "
//                  "TYPE_OF_CITY1, TYPE_OF_CITY2, CITY2, FSUBJ, DISTRICT, "
//                  "ADD, RAW, TYPE_OF_FSUBJ) "
//                  "VALUES(?, ?, ?, ?, ?, ?, ?, ?, ?, "
//                  "?, ?, ?, ?, ?, ?, ?, ?);");
//    query.prepare("INSERT INTO base1 (BUILD_ID, BUILD, STREET_ID, STREET,"
//                  "TYPE_OF_STREET, KORP, LITERA, CORRECT, CITY1, "
//                  "TYPE_OF_CITY1, TYPE_OF_CITY2, CITY2, FSUBJ, DISTRICT, "
//                  "ADD, RAW) "
//                  "VALUES(:bid, :b, :sid, :s, :tos, :k, :l, :c, :c1, "
//                  ":toc1, :toc2, :c2, :fs, :d, :a, :raw);");
//    query.prepare("INSERT INTO :table (:build_id, :build, :street_id, :street,"
//                  ":type_of_street, :korp, :litera, :correct, :city1, "
//                  ":type_of_city1, :type_of_city2, :city2, :fsubj, :district, "
//                  ":additional, :raw_addr) "
//                  "VALUES(:bid, :b, :sid, :s, :tos, :k, :l, :c, :c1, "
//                  ":toc1, :toc2, :c2, :fs, :d, :a, :raw);");
//    query.bindValue(":table", TableName);
//    query.bindValue(":build_id", MapColumnNames[BUILD_ID]);
//    query.bindValue(":street", MapColumnNames[STREET]);
//    query.bindValue(":street_id", MapColumnNames[STREET_ID]);
//    query.bindValue(":korp", MapColumnNames[KORP]);
//    query.bindValue(":build", MapColumnNames[BUILD]);
//    query.bindValue(":type_of_street", MapColumnNames[TYPE_OF_STREET]);
//    query.bindValue(":additional", MapColumnNames[ADDITIONAL]);
//    query.bindValue(":type_of_city1", MapColumnNames[TYPE_OF_CITY1]);
//    query.bindValue(":city1", MapColumnNames[CITY1]);
//    query.bindValue(":type_of_city2", MapColumnNames[TYPE_OF_CITY2]);
//    query.bindValue(":city2", MapColumnNames[CITY2]);
//    query.bindValue(":district", MapColumnNames[DISTRICT]);
//    query.bindValue(":fsubj", MapColumnNames[FSUBJ]);
//    query.bindValue(":raw_addr", MapColumnNames[RAW_ADDR]);
//    query.bindValue(":litera", MapColumnNames[LITERA]);
//    query.bindValue(":correct", MapColumnNames[CORRECT]);

//    query.bindValue(":bid",QString::number(a.getBuildId()));
//    query.bindValue(":s",a.getStreet());
//    query.bindValue(":sid",QString::number(a.getStreetId()));
//    query.bindValue(":k",a.getKorp());
//    query.bindValue(":b",a.getBuild());
//    query.bindValue(":tos",a.getTypeOfStreet());
//    query.bindValue(":a",a.getAdditional());
//    query.bindValue(":toc1",a.getTypeOfCity1());
//    query.bindValue(":c1",a.getCity1());
//    query.bindValue(":toc2",a.getTypeOfCity2());
//    query.bindValue(":c2",a.getCity2());
//    query.bindValue(":d",a.getDistrict());
//    query.bindValue(":l",a.getLitera());
//    query.bindValue(":fs",a.getFsubj());
//    query.bindValue(":c",QString(a.isCorrect()?"0":"1"));
//    query.bindValue(":raw",a.getRawAddress().join("\";\"").insert(0,'\"').append('\"'));
    QSqlQuery query;

    QString strF =
          "INSERT INTO  %1 (BUILD_ID, BUILD, STREET_ID, STREET,"
          "TYPE_OF_STREET, KORP, LITERA, CORRECT, CITY1, "
          "TYPE_OF_CITY1, TYPE_OF_CITY2, CITY2, FSUBJ, DISTRICT, "
          "ADDITIONAL, RAW, TYPE_OF_FSUBJ) "
          "VALUES('%2', '%3', '%4', '%5', '%6', "
          "'%7', '%8', '%9', '%10', '%11', '%12', '%13', "
          "'%14', '%15', '%16', '%17', '%18');";

    QString str =
            strF
            .arg(TableName)
            .arg(QString::number(a.getBuildId()))
            .arg(a.getBuild())
            .arg(QString::number(a.getStreetId()))
            .arg(a.getStreet())
            .arg(a.getTypeOfStreet())
            .arg(a.getKorp())
            .arg(a.getLitera())
            .arg(a.isCorrect()?"1":"0")
            .arg(a.getCity1())
            .arg(a.getTypeOfCity1())
            .arg(a.getTypeOfCity2())
            .arg(a.getCity2())
            .arg(a.getFsubj())
            .arg(a.getDistrict())
            .arg(a.getAdditional())
            .arg(a.getRawAddress().join("\";\"").insert(0,'\"').append('\"'))
            .arg(a.getTypeOfFSubjInString());

    if (!query.exec(str))
    {
        toDebug(objectName(),
                "Unable to make insert operation:"
                +a.toString(RAW)
                +":\r\n"+query.lastError().text());
//        assert(0);
    }
//    else
//        toDebug(objectName(),
//                "Success make insert operation");
}

void Database::selectAddress(QString sheet, int nRow, Address a)
{
    QString str = QString("SELECT STREET_ID, BUILD_ID, RAW "
                          "FROM %1 "
                          "WHERE TYPE_OF_FSUBJ = '%2' "
                          "  AND FSUBJ = '%3' "
                          "  AND DISTRICT = '%4' "
                          "  AND TYPE_OF_CITY1 = '%5' "
                          "  AND CITY1 = '%6' "
                          "  AND TYPE_OF_CITY2 = '%7' "
                          "  AND CITY2 = '%8'"
                          "  AND TYPE_OF_STREET = '%9' "
                          "  AND STREET = '%10' "
                          "  AND BUILD = '%11' "
                          "  AND KORP = '%12' "
                          "  AND LITERA = '%13' "
                          "  AND CORRECT = '1'; "
                          )
                  .arg(TableName)
                  .arg(a.getTypeOfFSubjInString())
                  .arg(a.getFsubj())
                  .arg(a.getDistrict())
                  .arg(a.getTypeOfCity1())
                  .arg(a.getCity1())
                  .arg(a.getTypeOfCity2())
                  .arg(a.getCity2())
                  .arg(a.getTypeOfStreet())
                  .arg(a.getStreet())
                  .arg(a.getBuild())
                  .arg(a.getKorp())
                  .arg(a.getLitera());
//    qDebug().noquote() << str;
    QSqlQuery query;
    if (!query.exec(str))
    {
        qDebug() << "Unable to execute query - exiting"
                 << endl
                 << query.lastError().text();
        return;
    }

    //Reading of the data
    QSqlRecord rec     = query.record();
    bool isExists=false;
    while (query.next()) {
        isExists=true;
        a.setRawAddress( query.value(rec.indexOf("RAW")).toString() );
        a.setBuildId( query.value(rec.indexOf("BUILD_ID")).toULongLong() );
        a.setStreetId( query.value(rec.indexOf("STREET_ID")).toULongLong() );
        emit addressFounded(sheet, nRow, a);
    }
    if(!isExists)
        emit addressNotFounded(sheet, nRow, a);
}

void Database::selectAddress(Address &a)
{
//    int n = qrand();
//    if(a.getBuildId()!=0 && a.getStreetId()!=0)
//        return;
//    if(           n%5==0
//               || n%5==1
//               || n%5==2
//               || n%5==3
//               /*|| n%5==4*/ )
//    {
//        a.setBuildId( 100000+qrand()%100000 );
//        a.setStreetId( 100000+qrand()%100000 );
//        return;
//    }
//    return;

    QSqlQuery query;
    if (!query.exec(QString("SELECT STREET_ID, BUILD_ID "
                    "FROM BASE"
                    "WHERE STREET = '%1'"
                    "  AND TYPE_OF_STREET = '%2.'"
                    "  AND CITY1 = '%3'"
                    "  AND BUILD = '%4'"
                    "  AND KORP = '%5';")
                    .arg(a.getStreet())
                    .arg(a.getTypeOfStreet())
                    .arg(a.getCity1())
                    .arg(a.getBuild())
                    .arg(a.getKorp()))) {
        qDebug() << "Unable to execute query - exiting"
                 << endl
                 << query.lastError().text();
        return;
    }

    //Reading of the data
    QSqlRecord rec     = query.record();
    while (query.next()) {
        a.setBuildId( query.value(rec.indexOf("BUILD_ID")).toULongLong() );
        a.setStreetId( query.value(rec.indexOf("STREET_ID")).toULongLong() );
    }
}

void Database::dropTable()
{
    qDebug().noquote() << "Database dropTable" << QThread::currentThreadId();
    QSqlQuery query;
//    query.prepare("DROP TABLE IF EXISTS :table;");
//    query.bindValue(":table", TableName);

    QString str =
            "DROP TABLE IF EXISTS base1;";

    if(!query.exec(str))
    {
        qDebug().noquote() << "dropTable error" << QThread::currentThreadId();

        toDebug(objectName(),
                "Unable to drop a table:\n"+query.lastError().text());
    }
    else
    {
        qDebug().noquote() << "dropTable success" << QThread::currentThreadId();

        toDebug(objectName(),
                "Success drop a table");
    }
//    if(_model!=nullptr)
//    {
//        delete _model;
//        _model=nullptr;
//    }
}
