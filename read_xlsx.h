#ifndef READ_XLSX_H
#define READ_XLSX_H

#include <QMainWindow>
#include <QtXlsx>
#include <clocale>

class Read_xlsx
{
private:
    QXlsx::Document _xls;
    QList<QString>  _buy;
    QList<QString>  _sell;
    QList<QString>  _date;
    QList<QString>  _commentary;
    QList<QString>  _nds;
    QList<QString>  _type;
    int             _ch_m;
    int             _count_s;
    int             _count_r_20;
public:
    Read_xlsx(QXlsx::Document &xlsx);

    int             _getGr_m(QString month);
    int             _getCh_m(QString month);
    int             _getCount_r(QString month);
    int             _getCount_s(QString month);
    QString         _getSalary(QString month);
    QStringList     _getAllMonth();
};

#endif // READ_XLSX_H
