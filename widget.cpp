#include "widget.h"
#include "ui_widget.h"

Widget::Widget(QWidget *parent) :
    QWidget(parent),
    ui(new Ui::Widget)
{
    ui->setupUi(this);
    _exc=new Excel;
    _db=new Database;
    _db->setBaseName("Base2.db");
    ui->lineEdit->setText("Base2.db");
    connect(_db, SIGNAL(toDebug(QString,QString)),
            _exc, SLOT(onDebug(QString,QString)));
    connect(_exc, SIGNAL(findRowInBase(QString,int,Address)),
            _db, SLOT(selectAddress(QString,int,Address)));
    connect(_db, SIGNAL(addressFounded(QString,int,Address)),
            _exc, SLOT(onFounedAddress(QString,int,Address)));
    connect(_db, SIGNAL(addressNotFounded(QString,int,Address)),
            _exc, SLOT(onNotFounedAddress(QString,int,Address)));

    ui->pushButton_2->setEnabled(false);
}

Widget::~Widget()
{
    delete ui;
    delete _exc;
    delete _db;
}

void Widget::open()
{
    QString str =
            QFileDialog::getOpenFileName(this,
                                         trUtf8("Укажите исходный файл"),
                                         "",
                                         tr("Excel (*.xls *.xlsx *.csv)"));
    if(str.isEmpty())
        return;
    if(str.contains(".csv"))
        _exc->runThreadOpenCsv(str);
    else
        _exc->runThreadOpen(str);
}

void Widget::on_pushButton_clicked()
{
    open();
    ui->pushButton_2->setEnabled(true);
}

void Widget::on_pushButton_2_clicked()
{
    _db->openBase();
    _exc->search();
    _exc->closeLog();
}

void Widget::on_pushButton_3_clicked()
{
    QString str =
            QFileDialog::getOpenFileName(this,
                                         trUtf8("Укажите файл БД"),
                                         "",
                                         tr("SQLite3 (*.db)"));
    if(str.isEmpty())
        return;
    ui->lineEdit->setText(str);
    _db->setBaseName(str);
}
