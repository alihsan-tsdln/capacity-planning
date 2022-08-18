#include "mainwindow.h"
#include "ui_mainwindow.h"

MainWindow::MainWindow(QWidget *parent)
    : QMainWindow(parent)
    , ui(new Ui::MainWindow)
{
    ui->setupUi(this);
    Init();
}

MainWindow::~MainWindow()
{
    delete ui;
}

int MainWindow::siparis_sayisi;
int MainWindow::siparis_edilen;
QStringList *lstMain;
QStringList *alph;

void MainWindow::Init()
{
    qsql.setDatabaseName("./db.db");

    if(qsql.open())
        ui->statusbar->showMessage("Veritabanına Bağlanıldı");

    else
        ui->statusbar->showMessage("Veritabanına Bağlanılamadı");


    QSqlQuery query;
    query.exec("select * from siparis");
    while(query.next())
        siparis_sayisi = query.value(1).toInt();

    //ui->siparis_sayisi_lbl->setText("Sipariş Sayısı : " + QString::number(siparis_sayisi));
    query.exec("select count(proje_adi) from proje");
    while(query.next())
    {
        siparis_edilen = query.value(0).toInt();
        /*ui->siparis_edilen_lbl->setText("Sipariş Edilen : " + QString::number(siparis_edilen));
        ui->kalan_siparis_lbl->setText ("Kalan Sipariş  : " + QString::number(100 * siparis_edilen / siparis_sayisi) + "%");
        */
    }

    lstMain = new QStringList();
    lstMain->append("proje_adi");
    lstMain->append("bohrwerk");
    lstMain->append("torna");
    lstMain->append("cnc");
    lstMain->append("kaynak");
    lstMain->append("montaj");
    lstMain->append("elektrik_montaj");

    alph = new QStringList();
    alph->append("A");
    alph->append("B");
    alph->append("C");
    alph->append("D");
    alph->append("E");
    alph->append("F");
    alph->append("G");
}

void MainWindow::on_girdi_btn_clicked()
{
    Makine_Girdisi *mg = new Makine_Girdisi();
    mg->show();
}


void MainWindow::on_ekle_btn_clicked()
{
    if(ui->ad_ln->text().isEmpty())
    {
        QMessageBox::information(this, "Proje Adı Boş", "Lütfen Proje Adını Giriniz");
        return;
    }


    QSqlQuery query;
    query.prepare("insert into proje values(:proje_adi, :bohrwerk, :torna, :cnc, :kaynak, :montaj, :elektrik_montaj)");
    query.bindValue(":proje_adi", ui->ad_ln->text());
    query.bindValue(":bohrwerk", ui->bohrwerk_ds->value());
    query.bindValue(":torna", ui->torna_ds->value());
    query.bindValue(":cnc", ui->cnc_ds->value());
    query.bindValue(":kaynak", ui->kaynak_ds->value());
    query.bindValue(":montaj", ui->montaj_ds->value());
    query.bindValue(":elektrik_montaj", ui->elektro_ds->value());
    query.exec();

    query.prepare("insert into proje_old values(:proje_adi, :bohrwerk, :torna, :cnc, :kaynak, :montaj, :elektrik_montaj)");
    query.bindValue(":proje_adi", ui->ad_ln->text());
    query.bindValue(":bohrwerk", ui->bohrwerk_ds->value());
    query.bindValue(":torna", ui->torna_ds->value());
    query.bindValue(":cnc", ui->cnc_ds->value());
    query.bindValue(":kaynak", ui->kaynak_ds->value());
    query.bindValue(":montaj", ui->montaj_ds->value());
    query.bindValue(":elektrik_montaj", ui->elektro_ds->value());
    query.exec();

    if(query.lastError().nativeErrorCode() == "1555")
    {
        QMessageBox::StandardButton reply;

        reply = QMessageBox::question(this, "Bilgilendirme", "Bu projeyi kayıt etmişsiniz. Verileri değiştirmek mi istiyorsunuz?", QMessageBox::Yes | QMessageBox::No);
        if(reply == QMessageBox::Yes)
        {
            query.prepare("update proje set bohrwerk = "
                          + QString::number(ui->bohrwerk_ds->value())
                          + ", torna = " + QString::number(ui->torna_ds->value())
                          + ", cnc = " + QString::number(ui->cnc_ds->value())
                          + ", kaynak = " + QString::number(ui->kaynak_ds->value())
                          + ", montaj = " + QString::number(ui->montaj_ds->value())
                          + ", elektrik_montaj = " + QString::number(ui->elektro_ds->value())
                          + " where proje_adi = '" + ui->ad_ln->text() + "'");

            if(!query.exec())
            {
                QMessageBox::critical(this, "Error", query.lastError().text());
            }

            query.prepare("update proje_old set bohrwerk = "
                          + QString::number(ui->bohrwerk_ds->value())
                          + ", torna = " + QString::number(ui->torna_ds->value())
                          + ", cnc = " + QString::number(ui->cnc_ds->value())
                          + ", kaynak = " + QString::number(ui->kaynak_ds->value())
                          + ", montaj = " + QString::number(ui->montaj_ds->value())
                          + ", elektrik_montaj = " + QString::number(ui->elektro_ds->value())
                          + " where proje_adi = '" + ui->ad_ln->text() + "'");

            if(!query.exec())
            {
                QMessageBox::critical(this, "Error", query.lastError().text());
            }
        }
    }

    query.exec("update siparis set siparis_edilen = " + QString::number(++siparis_sayisi) + " where id = 0");

    //ui->siparis_sayisi_lbl->setText("Sipariş Sayısı : " + QString::number(siparis_sayisi));
    query.exec("select count(proje_adi) from proje");
    while(query.next())
    {
        siparis_edilen = query.value(0).toInt();
    }
}


void MainWindow::on_goster_btn_clicked()
{
    Projeler *p = new Projeler();
    p->show();
}


void MainWindow::on_excel_ekle_btn_clicked()
{
    QString filepath = QFileDialog::getOpenFileName(this,tr("Excel Ekle"), ".", tr( "Microsoft Office Excel 365 (*.xlsx)"));

    if(!filepath.isEmpty())
    {
        auto excel     = new QAxObject("Excel.Application");
        excel->dynamicCall("SetVisible (bool Visible)", false); //Do not display the form
        excel->setProperty("DisplayAlerts" , false);
        auto workbooks = excel->querySubObject("Workbooks");
        auto workbook  = workbooks->querySubObject("Open(const QString&)", filepath);
        auto sheets    = workbook->querySubObject("Worksheets");
        auto sheet     = sheets->querySubObject("Item(int)", 1);

        int r = 1;
        QAxObject *cell;
        QSqlQuery query;
        QSqlQuery qry;

        while(true)
        {
            QString X="A"+QString::number(++r);
            cell = sheet->querySubObject("Range(QVariant, QVariant)",X);
            if(cell->dynamicCall("Value()").toString().isEmpty())
                break;
            qDebug() << lstMain->at(0);
            qDebug() << cell->dynamicCall("Value()").toString();
            query.prepare("insert into proje values(:proje_adi, :bohrwerk, :torna, :cnc, :kaynak, :montaj, :elektrik_montaj)");
            qry.prepare("insert into proje_old values(:proje_adi, :bohrwerk, :torna, :cnc, :kaynak, :montaj, :elektrik_montaj)");
            query.bindValue(":" + lstMain->at(0), cell->dynamicCall("Value()").toString());
            qry.bindValue(":" + lstMain->at(0), cell->dynamicCall("Value()").toString());



            for(int c = 2; c < 8; c++)
            {
                qDebug() << lstMain->at(c - 1);
                auto cCell = sheet->querySubObject("Cells(int,int)",r,c);
                query.bindValue(":" + lstMain->at(c - 1), cCell->dynamicCall("Value()").toDouble());
                qry.bindValue(":" + lstMain->at(c - 1), cCell->dynamicCall("Value()").toDouble());
                qDebug() << cCell->dynamicCall("Value()").toString();
            }

            query.exec();
            qry.exec();

            if(query.lastError().nativeErrorCode() == "1555")
            {
                continue;
            }

            query.exec("update siparis set siparis_edilen = " + QString::number(++siparis_sayisi) + " where id = 0");

            //ui->siparis_sayisi_lbl->setText("Sipariş Sayısı : " + QString::number(siparis_sayisi));
            query.exec("select count(proje_adi) from proje");
            while(query.next())
            {
                siparis_edilen = query.value(0).toInt();
            }

        }
    }
}

