#include "makine_girdisi.h"
#include "ui_makine_girdisi.h"

Makine_Girdisi::Makine_Girdisi(QWidget *parent) :
    QDialog(parent),
    ui(new Ui::Makine_Girdisi)
{
    ui->setupUi(this);
    Init();
}

Makine_Girdisi::~Makine_Girdisi()
{
    delete ui;
}

void Makine_Girdisi::on_sayi_spn_valueChanged(int arg1)
{
    int calisma_suresi = ui->sure_spn->value();

    ui->sonuc_lbl->setText("Toplam Çalışma Süresi : " + QString::number(calisma_suresi * arg1) + " dk");
}


void Makine_Girdisi::on_sure_spn_valueChanged(int arg1)
{
    int makine_sayisi = ui->sayi_spn->value();

    ui->sonuc_lbl->setText("Toplam Çalışma Süresi : " + QString::number(makine_sayisi * arg1) + " dk");
}


void Makine_Girdisi::on_ekle_btn_clicked()
{
    QSqlQuery query;
    query.prepare("insert into makine_suresi values (:makine_ismi, :makine_sayisi, :makine_suresi)");
    query.bindValue(":makine_ismi", ui->tip_cmb->currentText());
    query.bindValue(":makine_sayisi", ui->sayi_spn->value());
    query.bindValue(":makine_suresi", ui->sure_spn->value());
    query.exec();
    qDebug() << query.lastError().nativeErrorCode();

    if(query.lastError().nativeErrorCode() == "2067")
    {
        QMessageBox::StandardButton reply;
        reply = QMessageBox::question(this, "Bilgilendirme", "Makinenin değerini değiştirmek üzeresiniz. Kabul ediyor musunuz?",QMessageBox::Yes | QMessageBox::No);
        if(reply == QMessageBox::Yes)
            query.exec("update makine_suresi set makine_sayisi = " + QString::number(ui->sayi_spn->value()) + ", makine_suresi = " + QString::number(ui->sure_spn->value()) +" where makine_ismi = '" + ui->tip_cmb->currentText() + "'");

        else
            return;
    }
}


void Makine_Girdisi::on_tip_cmb_currentTextChanged(const QString &arg1)
{
    QSqlQuery query;
    query.exec("select makine_sayisi, makine_suresi from makine_suresi where makine_ismi = '" + arg1 + "'");
    query.next();
    ui->sayi_spn->setValue(query.value(0).toInt());
    ui->sure_spn->setValue(query.value(1).toInt());
}

void Makine_Girdisi::Init()
{
    QSqlQuery query;
    query.exec("select makine_sayisi, makine_suresi from makine_suresi where makine_ismi = '" + ui->tip_cmb->currentText() + "'");
    query.next();
    ui->sayi_spn->setValue(query.value(0).toInt());
    ui->sure_spn->setValue(query.value(1).toInt());
}

