#include "read_xlsx.h"
#include <iostream>

Read_xlsx::Read_xlsx(QXlsx::Document &xlsx)
{
    setlocale(LC_ALL, "Russian");

    int         row = 2;
    QXlsx::Cell *cell = xlsx.cellAt(row, 1);

    while (cell)
    {
        _date.push_back(QDate::longMonthName(xlsx.cellAt(row, 16)->dateTime().date().month(), QDate::StandaloneFormat));
        _buy.push_back(xlsx.cellAt(row, 23)->value().toString());
        _commentary.push_back(xlsx.cellAt(row, 28)->value().toString());
        _sell.push_back(xlsx.cellAt(row, 31)->value().toString());
        _nds.push_back(xlsx.cellAt(row, 32)->value().toString());
        _type.push_back(xlsx.cellAt(row, 38)->value().toString());
        row++;
        cell = xlsx.cellAt(row, 1);
    }
    _count_r_20 = 0;
    _ch_m = 0;
    _count_s = 0;
}


int             Read_xlsx::_getGr_m(QString month)
{
    int sum = 0;
    float nbr1, nbr2;
    bool  stat1, stat2;

    if (month == "Все")
    {
        for(int i = 0; i < _commentary.size(); i++)
        {
            if (_commentary[i] == "Закрыто" || _commentary[i] == "закрыто")
            {
                nbr1 = _buy[i].toFloat(&stat1);
                nbr2 = _sell[i].toFloat(&stat2);
                if (stat1 && stat2)
                    sum += (nbr1 - nbr2);
            }
        }
    }
    else
    {
        for(int i = 0; i < _commentary.size(); i++)
        {
            if ((_commentary[i] == "Закрыто" || _commentary[i] == "закрыто") && _date[i] == month)
            {
                nbr1 = _buy[i].toFloat(&stat1);
                nbr2 = _sell[i].toFloat(&stat2);
                if (stat1 && stat2)
                    sum += (nbr1 - nbr2);
            }
        }
    }
    return sum;
}

int             Read_xlsx::_getCh_m(QString month)
{
    int sum = 0;
    float nds = 1.2;
    float nbr1, nbr2;
    bool  stat1, stat2;

    if (month == "Все")
    {
        for(int i = 0; i < _commentary.size(); i++)
        {
            if (_commentary[i] == "Закрыто" || _commentary[i] == "закрыто")
            {
                nbr1 = _buy[i].toFloat(&stat1);
                nbr2 = _sell[i].toFloat(&stat2);
                if (stat1 && stat2)
                {
                    if (_nds[i] == "С НДС")
                        sum += ((nbr1 - nbr2) - ((nbr1 - nbr2) -  (nbr1 - nbr2) / nds));
                    else
                        sum += ((nbr1 - nbr2) - (nbr1 - nbr1 / nds));
                }
            }
        }
    }
    else
    {
        for(int i = 0; i < _commentary.size(); i++)
        {
            if ((_commentary[i] == "Закрыто" || _commentary[i] == "закрыто") && _date[i] == month)
            {
                nbr1 = _buy[i].toFloat(&stat1);
                nbr2 = _sell[i].toFloat(&stat2);
                if (stat1 && stat2)
                {
                    if (_nds[i] == "С НДС")
                        sum += ((nbr1 - nbr2) - ((nbr1 - nbr2) -  (nbr1 - nbr2) / nds));
                    else
                        sum += ((nbr1 - nbr2) - (nbr1 - nbr1 / nds));
                }
            }
        }
    }
    _ch_m = sum;
    return sum;
}

int             Read_xlsx::_getCount_r(QString month)
{
    int count = 0;

    if (month == "Все")
    {
        for(int i = 0; i < _commentary.size(); i++)
        {
            if (_commentary[i] == "Закрыто" || _commentary[i] == "закрыто")
            {
                if (_type[i] == "20т")
                    _count_r_20++;
                count++;
            }
        }
    }
    else
    {
        for(int i = 0; i < _commentary.size(); i++)
        {
            if ((_commentary[i] == "Закрыто" || _commentary[i] == "закрыто") && _date[i] == month)
            {
                if (_type[i] == "20т")
                    _count_r_20++;
                count++;
            }
        }
    }
    return count;
}

int             Read_xlsx::_getCount_s(QString month)
{
    int count = 0;

    if (month == "Все")
    {
        for(int i = 0; i < _commentary.size(); i++)
        {
            if (_commentary[i] == "Срыв" || _commentary[i] == "срыв")
            {
                count++;
            }
        }
    }
    else
    {
        for(int i = 0; i < _commentary.size(); i++)
        {
            if ((_commentary[i] == "Срыв" || _commentary[i] == "срыв") && _date[i] == month)
            {
                count++;
            }
        }
    }
    _count_s = count;
    return count;
}

QString             Read_xlsx::_getSalary(QString month)
{
    QString ret;
    int salary = 30000;
    float i = 0.025;
    float ndfl = 1.145;
    int all_salary = 882709;

    salary += _ch_m * i;

    if (_count_r_20 / 4 < 70 && _count_s > 0)
    {
        int less = ((70 - _count_r_20 / 4) <= _count_s) ? (70 - _count_r_20 / 4) : _count_s;
        std::cout << _count_r_20 / 4 <<" | " << less << std::endl;
        salary -= less * 1000;
    }
    if (salary < 30000)
        salary = 30000;
    ret = QString::number(salary);
    all_salary += (salary - 30000) * 4 * ndfl;
    if (_ch_m - all_salary - 500000 > 0)
    {
        int temp = (_ch_m - all_salary - 500000) / 100 * 10;
        ret += " | " + QString::number(temp);
    }
    return ret;
}

QStringList  Read_xlsx::_getAllMonth()
{
    QStringList all_month;
    bool flag = true;

    for(int i = 0; i < _date.size(); i++)
    {
        for (int j = 0; j < all_month.size(); j++)
        {
            flag = true;
            if (_date[i] == all_month[j])
            {
                flag = false;
                break;
            }
        }
        if (flag == true)
            all_month.push_back(_date[i]);
    }
    return all_month;
}
