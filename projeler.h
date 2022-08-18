#ifndef PROJELER_H
#define PROJELER_H

#include <QDialog>
#include <QSqlQuery>
#include <QSqlQueryModel>
#include <QSqlError>
#include "makine_girdisi.h"

namespace Ui {
class Projeler;
}

class Projeler : public QDialog
{
    Q_OBJECT

public:
    explicit Projeler(QWidget *parent = nullptr);
    ~Projeler();

private slots:
    void on_sil_btn_clicked();

    void on_projeler_table_cellClicked(int row, int column);

    void on_degistir_btn_clicked();

    void on_excel_btn_clicked();

private:
    Ui::Projeler *ui;
    void Init();
    QSqlQueryModel *model;
    QSqlQueryModel *modelToplam;
    void guncelle();
};

#endif // PROJELER_H
