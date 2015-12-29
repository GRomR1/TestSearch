#ifndef WIDGET_H
#define WIDGET_H

#include <QWidget>
#include "excel.h"
#include "database.h"

namespace Ui {
class Widget;
}

class Widget : public QWidget
{
    Q_OBJECT

public:
    explicit Widget(QWidget *parent = 0);
    ~Widget();

private slots:
    void on_pushButton_clicked();

    void on_pushButton_2_clicked();

    void on_pushButton_3_clicked();

private:
    Ui::Widget *ui;
    Excel *_exc;
    Database *_db;

    void open();
};

#endif // WIDGET_H