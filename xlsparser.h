#ifndef XLSPARSER_H
#define XLSPARSER_H

#include <QObject>
#include "defines.h"
#include "address.h"

class XlsParser : public QObject
{
    Q_OBJECT
public:
    explicit XlsParser(QObject *parent = 0);
    ~XlsParser();

signals:
    void finished();
    void sheetParsed(QString sheet);
    void rowParsed(QString sheet, int row, Address addr);

public slots:
    void process();
    void onReadHead(const QString sheet,
                    MapAddressElementPosition head);
    void onReadRow(const QString &sheet,
                   const int &rowNumber,
                   const QStringList &row);
    void onIsOneColumn(bool b);

private:
    QMap<QString, MapAddressElementPosition> _mapHead;
    QMap<QString, MapAddressElementPosition> _mapPHead;
    bool _error;
    bool _isOneColumn;

    bool parseObject(QString &str, QString &result, QString regPattern, int regCap);
    bool parseObject(QString &str, QStringList &result, QString regPattern);
};

#endif // XLSPARSER_H
