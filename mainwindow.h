#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include <QMainWindow>
#include <QtXlsx>
#include "read_xlsx.h"

QT_BEGIN_NAMESPACE
namespace Ui { class MainWindow; }
QT_END_NAMESPACE

class MainWindow : public QMainWindow
{
    Q_OBJECT

public:
    MainWindow(QWidget *parent = nullptr);
    ~MainWindow();

    void   Change(const QString &arg1);

private slots:

    void on_pushButton_clicked();

private:
    Ui::MainWindow *ui;

    Read_xlsx *_shablon;
};
#endif // MAINWINDOW_H
