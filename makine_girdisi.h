#ifndef MAKINE_GIRDISI_H
#define MAKINE_GIRDISI_H

#include <QDialog>
#include <QSql>
#include <QSqlDatabase>
#include <QSqlError>
#include <QSqlQuery>
#include <QSqlQueryModel>
#include <QMessageBox>
#include <QComboBox>

namespace Ui {
class Makine_Girdisi;
}

class Makine_Girdisi : public QDialog
{
    Q_OBJECT

public:
    explicit Makine_Girdisi(QWidget *parent = nullptr);
    ~Makine_Girdisi();


private slots:
    void on_sayi_spn_valueChanged(int arg1);

    void on_sure_spn_valueChanged(int arg1);

    void on_ekle_btn_clicked();

    void on_tip_cmb_currentTextChanged(const QString &arg1);

private:
    Ui::Makine_Girdisi *ui;

    void Init();
};

#endif // MAKINE_GIRDISI_H
