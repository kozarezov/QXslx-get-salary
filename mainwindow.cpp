#include "mainwindow.h"
#include "ui_mainwindow.h"
#include <string>
#include "read_xlsx.h"
#include <QLabel>
#include <iostream>

MainWindow::MainWindow(QWidget *parent)
    : QMainWindow(parent)
    , ui(new Ui::MainWindow)
{
   //connect(ui->comboBox, &QComboBox::activated, this, &MainWindow::Change);

    ui->setupUi(this);

}

MainWindow::~MainWindow()
{
    delete ui;
}


void MainWindow::on_pushButton_clicked()
{
    QString path = ui->textEdit->toPlainText();
    //QString path = "/Users/deniskozarezov/Desktop/шаблон.xlsx";
    //ui->textEdit->clear();
    QXlsx::Document xlsx(path);

    Read_xlsx shablon(xlsx);
    _shablon = &shablon;

    QString now = ui->comboBox->currentText();
    Change(now);
}

void   MainWindow::Change(const QString &arg1)
{
    ui->gr_m->setNum(_shablon->_getGr_m(arg1));
    ui->ch_m->setNum(_shablon->_getCh_m(arg1));
    ui->count_r->setNum(_shablon->_getCount_r(arg1));
    ui->count_s->setNum(_shablon->_getCount_s(arg1));
    ui->salary->setText(_shablon->_getSalary(arg1));
}

