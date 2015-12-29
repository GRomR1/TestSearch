#ifndef DATABASE_H
#define DATABASE_H

#include <QObject>
#include <QtSql>
#include "address.h"
#include "defines.h"

//const QString TableName = "base1";

typedef QList< Address > ListAddress;

class Database : public QObject
{
    Q_OBJECT
public:
    explicit Database(QObject *parent = 0);
    ~Database();
    QSqlTableModel *getModel();
    void createConnection();
    void dropTable();
    void createTable();

signals:
    void headReaded(QStringList head);
    void headParsed(MapAddressElementPosition head);
    void rowReaded(int rowNumber);
    void rowParsed(int rowNumber);
    void readedRows(int count);
    void countRows(int count);
    void workingWithOpenBase();
    void baseOpened();

    void toDebug(QString,QString);

    void addressFounded(QString sheet, int nRow, Address a);
    void addressNotFounded(QString sheet, int nRow, Address a);

public slots:
    void openBase();

    ListAddress search(QString sheetName, ListAddress addr);

    void insertAddress(Address a);

    void setBaseName(QString name);
    void clear();

    void selectAddress(QString sheet, int nRow, Address a);

private:
    QThread *_thread;
    QSqlTableModel *_model;
    QSqlDatabase _db;
    QString _baseName;
    bool _connected;

    void openTableToModel();
//    void insertAddress(const Address &a);
    void selectAddress(Address &a);
};

#endif // DATABASE_H
