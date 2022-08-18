#include "projeler.h"
#include "ui_projeler.h"
#include <QAbstractItemModel>
#include "mainwindow.h"

QStringList *lst;
QStringList *lstCol;
QStringList *alphabet;

Projeler::Projeler(QWidget *parent) :
    QDialog(parent),
    ui(new Ui::Projeler)
{
    ui->setupUi(this);

    lst = new QStringList();
    lst->append("Proje Adı");
    lst->append("Bohrwerk");
    lst->append("Torna");
    lst->append("CNC");
    lst->append("Kaynak");
    lst->append("Montaj");
    lst->append("Elektrik Montaj");

    lstCol = new QStringList();
    lstCol->append("proje_adi");
    lstCol->append("bohrwerk");
    lstCol->append("torna");
    lstCol->append("cnc");
    lstCol->append("kaynak");
    lstCol->append("montaj");
    lstCol->append("elektrik_montaj");

    alphabet = new QStringList();
    alphabet->append("A");
    alphabet->append("B");
    alphabet->append("C");
    alphabet->append("D");
    alphabet->append("E");
    alphabet->append("F");
    alphabet->append("G");

    for(int i = 0; i < 3; i++)
    {
        ui->sonuc_table->insertRow(i);
    }

    ui->projeler_table->insertColumn(0);

    for(int i = 0; i < 6; i++)
    {
        ui->sonuc_table->insertColumn(i);
        ui->projeler_table->insertColumn(i + 1);
    }

    ui->projeler_table->setHorizontalHeaderLabels(*lst);
    lst->removeAt(0);
    ui->sonuc_table->setHorizontalHeaderLabels(*lst);

    QStringList *lstV = new QStringList();
    lstV->append("TOPLAM İŞLENMEYENLER");
    lstV->append("KAPASİTE");
    lstV->append("TOPLAM GÜN");

    ui->sonuc_table->setVerticalHeaderLabels(*lstV);

    Init();
}

Projeler::~Projeler()
{
    delete ui;
}

QString proje_adi;

void Projeler::Init()
{
    guncelle();

    ui->siparis_sayisi_lbl->setText("Sipariş Sayısı : " + QString::number(MainWindow::siparis_sayisi));
    ui->siparis_edilen_lbl->setText("Sipariş Edilen : " + QString::number(MainWindow::siparis_edilen));
    ui->kalan_siparis_lbl->setText ("Kalan Sipariş  : " + QString::number(100 * (double)MainWindow::siparis_edilen / MainWindow::siparis_sayisi, 'f', 2) + "%");

    QSqlQuery qry;

    qry.exec("select sum(bohrwerk),"
                " sum(torna), sum(cnc), sum(kaynak), sum(montaj),"
                " sum(elektrik_montaj) from proje");
    qry.next();


    for(int i = 0; i < 6; i++)
    {
        ui->sonuc_table->setItem(0, i, new QTableWidgetItem(qry.value(i).toString()));
    }

    qry.exec("select * from makine_suresi");
    while(qry.next())
    {
        if(lst->contains(qry.value(0).toString()))
        {
            ui->sonuc_table->setItem(1, lst->indexOf(qry.value(0).toString()), new QTableWidgetItem(QString::number(qry.value(1).toDouble() * qry.value(2).toDouble())));
        }
    }

    for(int i = 0; i < 6; i++)
    {
        QTableWidgetItem *item(ui->sonuc_table->item(1, i));

        if(item)
        {
            ui->sonuc_table->setItem(2, i, new QTableWidgetItem(QString::number(ui->sonuc_table->item(0, i)->text().toDouble() / ui->sonuc_table->item(1, i)->text().toDouble(), 'f', 2)));
        }
    }
}

void Projeler::guncelle()
{
    bool key = true;
    while(ui->projeler_table->rowCount() > 0)
    {
        ui->projeler_table->removeRow(0);
    }

    QSqlQuery query;
    QSqlQuery qry;
    query.exec("select * from proje");
    qry.exec("select * from proje_old");

    while(query.next() && qry.next())
    {
        ui->projeler_table->insertRow(ui->projeler_table->rowCount());
        for(int i = 0; i < 7; i++)
        {
            ui->projeler_table->setItem(ui->projeler_table->rowCount() - 1, i, new QTableWidgetItem(qry.value(i).toString()));
            if(key)
                ui->projeler_table->item(ui->projeler_table->rowCount() - 1, i)->setBackground(QBrush(QColor::fromRgba(qRgba(102, 255, 0, 30))));
        }
        ui->projeler_table->insertRow(ui->projeler_table->rowCount());
        ui->projeler_table->setItem(ui->projeler_table->rowCount() - 1, 0, new QTableWidgetItem("Gerçekleşen Süre"));
        if(key)
            ui->projeler_table->item(ui->projeler_table->rowCount() - 1, 0)->setBackground(QBrush(QColor::fromRgba(qRgba(102, 255, 0, 30))));
        for(int i = 1; i < 7; i++)
        {
            ui->projeler_table->setItem(ui->projeler_table->rowCount() - 1, i, new QTableWidgetItem(QString::number(qry.value(i).toDouble() - query.value(i).toDouble())));
            if(key)
                ui->projeler_table->item(ui->projeler_table->rowCount() - 1, i)->setBackground(QBrush(QColor::fromRgba(qRgba(102, 255, 0, 30))));
                //ui->projeler_table->item(ui->projeler_table->rowCount() - 1, i)->setBackground(brush);
        }

        ui->projeler_table->insertRow(ui->projeler_table->rowCount());
        ui->projeler_table->setItem(ui->projeler_table->rowCount() - 1, 0, new QTableWidgetItem("Tamamlanan"));
        if(key)
            ui->projeler_table->item(ui->projeler_table->rowCount() - 1, 0)->setBackground(QBrush(QColor::fromRgba(qRgba(102, 255, 0, 30))));
        for(int i = 1; i < 7; i++)
        {
            ui->projeler_table->setItem(ui->projeler_table->rowCount() - 1, i, new QTableWidgetItem(QString::number(100 - (100 * query.value(i).toDouble() / qry.value(i).toDouble())) + "%"));
            if(key)
                ui->projeler_table->item(ui->projeler_table->rowCount() - 1, i)->setBackground(QBrush(QColor::fromRgba(qRgba(102, 255, 0, 30))));
                //ui->projeler_table->item(ui->projeler_table->rowCount() - 1, i)->setBackground(brush);
        }
        key = !key;
    }


}

int rowClick = NULL;

void Projeler::on_sil_btn_clicked()
{
    QSqlQuery query;
    query.exec("delete from proje where proje_adi = '" + proje_adi + "'");
    query.exec("delete from proje_old where proje_adi = '" + proje_adi + "'");

    if(rowClick)
    {
        for(int i = (rowClick / 3) * 3; i < ((rowClick / 3) * 3 + 3); i++)
        {
            ui->projeler_table->removeRow(i);
        }
    }

    query.exec("select count(proje_adi) from proje");
    query.next();
    MainWindow::siparis_edilen = query.value(0).toInt();

    Init();

}

void Projeler::on_projeler_table_cellClicked(int row, int column)
{
    rowClick = row;
    proje_adi = ui->projeler_table->item(row, 0)->text();
}

void Projeler::on_degistir_btn_clicked()
{
    QMessageBox::StandardButton reply;
    reply = QMessageBox::question(this, "Bilgilendirme", "Değerleri değiştirmek istediğinizden emin misiniz?", QMessageBox::Yes | QMessageBox::No);

    if(reply == QMessageBox::Yes)
    {
        QSqlQuery query;

        for(int i = 1; i < lstCol->count(); i++)
        {
            for(int j = 0; j < ui->projeler_table->rowCount(); j++)
            {
                query.exec("update proje_old set " + lstCol->at(i)
                           + " = " + ui->projeler_table->item(j,i)->text()
                           + " where proje_adi = '" + ui->projeler_table->item(j,0)->text() + "'");

                j++;

                //qDebug() << ui->projeler_table->item(j-1,i)->text().toDouble() - ui->projeler_table->item(j,i)->text().toDouble();

                if(ui->projeler_table->item(j-1,i)->text().toDouble() > ui->projeler_table->item(j,i)->text().toDouble())
                {
                    query.exec("update proje set " + lstCol->at(i)
                               + " = " + QString::number(ui->projeler_table->item(j-1,i)->text().toDouble() - ui->projeler_table->item(j,i)->text().toDouble())
                               + " where proje_adi = '" + ui->projeler_table->item(j-1,0)->text() + "'");
                }

                else
                {
                    query.exec("update proje set " + lstCol->at(i)
                               + " = 0  where proje_adi = '" + ui->projeler_table->item(j-1,0)->text() + "'");
                }


                j++;
            }
        }

        Init();
    }
}


void Projeler::on_excel_btn_clicked()
{
    QString filepath=QFileDialog::getSaveFileName( this ,tr("Save orbit"), "." ,tr( "Microsoft Office Excel 2007 (*.xlsx)" )); //Get the save path

    if (!filepath.isEmpty()){
        QAxObject *excel = new QAxObject(this);
        excel->setControl("Excel.Application"); //Connect to Excel control
        excel->dynamicCall("SetVisible (bool Visible)", false); //Do not display the form
        excel->setProperty("DisplayAlerts" , false); //Do not display any warning messages. If it is true, a prompt similar to "File has been modified, whether to save" will appear when closing
        QAxObject *workbooks = excel->querySubObject("WorkBooks"); //Get workbook collection
        workbooks->dynamicCall("Add"); //Create a new workbook
        QAxObject *workbook = excel->querySubObject("ActiveWorkBook"); //Get the current workbook
        QAxObject *worksheets = workbook->querySubObject("Sheets"); //Get worksheet collection
        QAxObject *worksheet = worksheets->querySubObject("Item(int)" ,1); //Get worksheet 1 of the worksheet collection, namely sheet1
        QAxObject *cellX;

        lst = new QStringList();
        lst->append("Proje Adı");
        lst->append("Bohrwerk");
        lst->append("Torna");
        lst->append("CNC");
        lst->append("Kaynak");
        lst->append("Montaj");
        lst->append("Elektrik Montaj");


        for(int i = 0; i < alphabet->count(); i++)
        {
            QString X = alphabet->at(i) + QString::number(1);
            cellX = worksheet->querySubObject("Range(QVariant, QVariant)",X);
            cellX->dynamicCall("SetValue(const QVariant&)",QVariant(lst->at(i)));
        }

        int rowC = ui->projeler_table->rowCount();

        for(int j = 0; j < rowC; j++)
        {
            for(int i = 0; i < alphabet->count(); i++)
            {
                QString X= alphabet->at(i) + QString::number(j+2);
                cellX = worksheet->querySubObject("Range(QVariant, QVariant)",X);
                qDebug() << QVariant(ui->projeler_table->item(j,i)->text());

                if(ui->projeler_table->item(j,i)->text().contains("."))
                {
                    QStringList a = ui->projeler_table->item(j,i)->text().split(".");
                    cellX->dynamicCall("SetValue(const QVariant&)", QVariant(a.at(0) + "," + a.at(1)));
                }

                else
                {
                    cellX->dynamicCall("SetValue(const QVariant&)", QVariant(ui->projeler_table->item(j,i)->text()));
                }

            }
        }

        workbook->dynamicCall("SaveAs(const QString&)",QDir::toNativeSeparators(filepath));//Save to filepath, be sure to use QDir::toNativeSeparators to convert the "/" in the path to "\", otherwise It must not be saved.
        workbook->dynamicCall("Close()");//Close the workbook
        excel->dynamicCall("Quit()");//Close excel
        delete excel;
        excel=NULL;


    }
}

