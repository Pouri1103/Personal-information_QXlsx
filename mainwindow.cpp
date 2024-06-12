#include "mainwindow.h"
#include "ui_mainwindow.h"
#include <QSqlDatabase>
#include <QSqlQuery>
#include <QSqlQueryModel>
#include <QMessageBox>
#include <QList>

#include "xlsxdocument.h"
using namespace QXlsx;

//----------------------------

Document xlsx("user.xlsx");
QString val[4];

//----------------------------

MainWindow::MainWindow(QWidget *parent)
    : QMainWindow(parent)
    , ui(new Ui::MainWindow)
{
    ui->setupUi(this);

    QSqlDatabase db = QSqlDatabase::addDatabase("QSQLITE");
    db.setDatabaseName("users.db");

    if (!db.open())
        QMessageBox::critical(this,"error DB","Could not open the SQLite database.");
    QSqlQuery("create table myxlsx (Name text, FamilyName Text, NationalCode Text unique, Call Text)");
}

/////////////////////////////////////////////////////////////////////

MainWindow::~MainWindow()
{
    delete ui;
}

/////////////////////////////////////////////////////////////////////

void MainWindow::on_pb_readDB_clicked()
{
    QSqlQueryModel *m = new QSqlQueryModel;
    QSqlQuery q("select * from myxlsx");
    m->setQuery(std::move(q));
    ui->tableView->setModel(m);
}


void MainWindow::on_pb_writeExcel_clicked()
{
    xlsx.mergeCells("A2:D2");
    xlsx.write("A2","Info Members");
    xlsx.write("A1","Numbers");
    xlsx.write("A3","Name");
    xlsx.write("B3","Family Name");
    xlsx.write("C3","NationalCode");
    xlsx.write("D3","Phone Number");

    QSqlQuery q("select * from myxlsx");
    int num = 0;
    while (q.next())
    {
        xlsx.write("A"+QString::number(num+4),q.value(0).toString());
        xlsx.write("B"+QString::number(num+4),q.value(1).toString());
        xlsx.write("C"+QString::number(num+4),q.value(2).toString());
        xlsx.write("D"+QString::number(num+4),q.value(3).toString());
        num++;
    }

    xlsx.write("B1",QString::number(num));

    xlsx.setColumnWidth(1,20);
    xlsx.setColumnWidth(2,20);
    xlsx.setColumnWidth(3,20);
    xlsx.setColumnWidth(4,20);

    if (xlsx.save())
        QMessageBox::information(this,"Ok save","Saved successfully");
    else
        QMessageBox::critical(this,"Error save","Not be saved !!!");

}


void MainWindow::on_pb_readExcel_clicked()
{
    int rows = xlsx.read("B1").toInt() + 3;

    for (int row=4;row<=rows;row++)
    {
        for (int col=1;col<=4;col++)
            val[col] = xlsx.read(row,col).toString();

        QSqlQuery("insert into myxlsx values('"+val[1]+"','"+val[2]+"','"+val[3]+"','"+val[4]+"')");
    }
    on_pb_readDB_clicked();
}


void MainWindow::on_pb_add_clicked()
{
    if (ui->le_name->text()!="" && ui->le_fName->text()!="" && ui->le_nCode->text()!="" && ui->le_phone->text()!="")
    {
        QSqlQuery("insert into myxlsx values('"+ui->le_name->text()+"','"+ui->le_fName->text()+"','"+ui->le_nCode->text()+"','"+ui->le_phone->text()+"')");
        ui->le_name->clear();
        ui->le_fName->clear();
        ui->le_nCode->clear();
        ui->le_phone->clear();
        on_pb_readDB_clicked();
    }
}

