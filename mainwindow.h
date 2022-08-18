#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include <QMainWindow>
#include <QSql>
#include <QSqlDatabase>
#include <QSqlError>
#include <QSqlQuery>
#include <QSqlQueryModel>
#include "makine_girdisi.h"
#include "projeler.h"
#include <QAxObject>
#include <QFileDialog>

QT_BEGIN_NAMESPACE
namespace Ui { class MainWindow; }
QT_END_NAMESPACE

class MainWindow : public QMainWindow
{
    Q_OBJECT

public:
    MainWindow(QWidget *parent = nullptr);
    ~MainWindow();
    static int siparis_sayisi;
    static int siparis_edilen;

private slots:
    void on_girdi_btn_clicked();

    void on_ekle_btn_clicked();

    void on_goster_btn_clicked();

    void on_excel_ekle_btn_clicked();

private:
    Ui::MainWindow *ui;
    QSqlDatabase qsql = qsql.addDatabase("QSQLITE");
    QSqlQueryModel *model;
    void Init();
};
#endif // MAINWINDOW_H
