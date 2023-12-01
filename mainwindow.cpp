#include "mainwindow.h"
#include "ui_mainwindow.h"
//#include "xlsxcellrange.h"
//#include "xlsxchart.h"
//#include "xlsxrichstring.h"
#include "xlsxworkbook.h"
//Подключаем библиотеку для Excel
#include "xlsxdocument.h"
//#include "xlsxchartsheet.h"
using namespace QXlsx;





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

//Открываем файл с компьютера через кнопку
//Читаем Excel файл
int sheetIndexNumber = 0;
void MainWindow::on_btnFOpen_clicked()
{
    filePath = QFileDialog::getOpenFileName(this, "Выбрать файл", " ", "*.xls, *.xlsx");//Задаем название диалогового окна и параметры открываемого файла
        if(filePath.isEmpty()) {//Проверяем на пустоту и выводим сообщение
            msgBox->setWindowTitle("ВНИМАНИЕ!");
            msgBox->setText("<b>Не выбран файл!</b>");
            msgBox->exec();
        }else{
            ui->textLabel->setText("Открыт файл: " + filePath);
            ui->statusbar->clearMessage();
            ui->statusbar->showMessage(filePath);
        }

        QXlsx::Document xlsx;
        Document doc(filePath);

        foreach(QString currentSheetName, doc.sheetNames()){
            AbstractSheet* currentSheet = doc.sheet(currentSheetName);
            if(NULL == currentSheet)
                continue;

        int maxRow = -1;
        int maxCol = -1;
        currentSheet->workbook()->setActiveSheet(sheetIndexNumber);
        Worksheet* wsheet = (Worksheet*) currentSheet->workbook()->activeSheet();
        if(NULL == wsheet)
            continue;

        QString strSheetName = wsheet->sheetName();
        qDebug() << strSheetName;

        QVector<CellLocation> clList = wsheet->getFullCells(&maxRow, &maxCol);
        QVector<QVector<QString>> cellValues;

        for (int rc = 0; rc < maxRow; rc++){
            QVector<QString> tempValue;
            for(int cc = 0; cc < maxCol; cc++){
                tempValue.push_back(QString(""));
            }
            cellValues.push_back(tempValue);
        }

        for (int ic = 0; ic < clList.size(); ++ic){
            CellLocation cl = clList.at(ic);

            int row = cl.row - 1;
            int col = cl.col - 1;

//            QSharedPointer<Cell> ptrCell = cl.cell;

            QVariant var = cl.cell->dateTime();// .data()->value();
            QString str = var.toString();

            cellValues[row][col] = str;
        }
        for(int rc = 0; rc < maxRow; rc++){
            for(int cc = 0; cc < maxCol; cc++){
                QString strCell = cellValues[rc][cc];
                qDebug() << "( row :" << rc
                         << ", col :" << cc
                         << ") " << strCell;
            }
        }


        }


}
//Закрыть приложение кнопкой
void MainWindow::on_btnClose_clicked(){close();}
