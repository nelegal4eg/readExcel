#include "mainwindow.h"
#include "ui_mainwindow.h"
//Подключаем библиотеку для Excel
#include "xlsxdocument.h"
//#include "xlsxchartsheet.h"
//#include "xlsxcellrange.h"
//#include "xlsxchart.h"
//#include "xlsxrichstring.h"
//#include "xlsxworkbook.h"
using namespace QXlsx;


#include <QFileInfo>
#include <QFile>

#include <QAxWidget>
#include <QAxObject>




MainWindow::MainWindow(QWidget *parent)
    : QMainWindow(parent)
    , ui(new Ui::MainWindow)
{
    ui->setupUi(this);
    //Устанавливаем название основного окна
    setWindowTitle("Считывание EXCEL файлов в базу");
    //Подключаемся к базе
    //db = QSqlDatabase::addDatabase("QSQLITE");
    //db.setDatabaseName("./../readExcelPeople/readexcel/readexcel.db");
    //Создаем объект окна сообщений
    msgBox = new QMessageBox(this);
    msgBox->setIcon(QMessageBox::Information);
    //Проверяем подключение и выводим сообщение
    //if (db.open()) {
        //ui->statusbar->showMessage("Вы успешно подключены к базе данных: " + db.databaseName());
        //Выводим таблицу в окно приложения
        /*modelT = new QSqlTableModel(this, db);
        modelT->setTable("personnels");
        modelT->select();
        ui->tableView->setModel(modelT);
    }else{
        ui->statusbar->showMessage("При подключении к базе произошла ошибка: " + db.lastError().databaseText());*/
    }
//}

MainWindow::~MainWindow()
{
    delete ui;
}
//Закрыть приложение в меню
void MainWindow::on_exit_triggered(){close();}

//Открываем файл с компьютера
//Читаем Excel файл
void MainWindow::on_btnFOpen_clicked()
{
    filePath = QFileDialog::getOpenFileName(this, "Выбрать файл", " ", "*.xls, *.xlsx");

        if(filePath.isEmpty()) {
            msgBox->setWindowTitle("ВНИМАНИЕ!");
            msgBox->setText("<b>Не выбран файл!</b>");
            msgBox->exec();
        }else{
            ui->textLabel->setText("Открыт файл: " + filePath);
            ui->statusbar->clearMessage();
            ui->statusbar->showMessage(filePath);
        }

        QXlsx::Document xlsx;
}
//Закрыть приложение кнопкой
void MainWindow::on_btnClose_clicked(){close();}
