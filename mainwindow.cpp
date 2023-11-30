#include "mainwindow.h"
#include "ui_mainwindow.h"


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

        QAxObject* excel = new QAxObject("Excel.Application", this);
        QAxObject* workbooks = excel->querySubObject("Workbooks");//рабочая книга
        QAxObject* workbook = workbooks->querySubObject("Open(const QString)", filePath);
        excel->dynamicCall("SetVisible(bool)", false);//видимость документа

        QAxObject* worksheet = workbook->querySubObject("Worksheets(int)", 1);

        //Получаем количество строк и столбцов
        QAxObject* usedrange = worksheet->querySubObject("UsedRange");
        QAxObject* rows = usedrange->querySubObject("Rows");
        QAxObject* columns = usedrange->querySubObject("Columns");

        int intRowStart = usedrange->property("Row").toInt();
        int intColStart = usedrange->property("Column").toInt();
        int intCols = columns->property("Count").toInt();
        int intRows = rows->property("Count").toInt();

        qDebug() << intRows;
        qDebug() << intCols;

        ui->tableWidget->setColumnCount(intColStart + intCols);
        ui->tableWidget->setRowCount(intRowStart + intRows);

        //Заполняем таблицу
        for(int row=0; row < intRows; row++){
            for(int col = 0; col < intCols; col++){
                QAxObject* cell = worksheet->querySubObject("Cells(int, int)", row + 1, col +1);
                QVariant value = cell->dynamicCall("Value()");
                QTableWidgetItem* item = new QTableWidgetItem(value.toString());
                ui->tableWidget->setItem(row, col, item);
            }
        }
        //Закрывем файл
        delete worksheet;
        workbook->dynamicCall("Close (Boolean)", false);
        delete workbooks;
        excel->dynamicCall("Quit (void)");
        delete excel;
}
//Закрыть приложение кнопкой
void MainWindow::on_btnClose_clicked(){close();}
