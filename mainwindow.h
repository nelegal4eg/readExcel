#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include <QMainWindow>
#include <QSqlDatabase>
#include <QSqlError>
#include <QSqlTableModel>
#include <QFileSystemModel>
#include <QFile>
#include <QFileDialog>
#include <QString>
#include <QTime>
#include <QMessageBox>
#include <QSharedPointer>



namespace Ui { class MainWindow; }


class MainWindow : public QMainWindow
{
    Q_OBJECT

public:
    MainWindow(QWidget *parent = nullptr);
    ~MainWindow();

private slots:
    void on_exit_triggered();    

    void on_btnFOpen_clicked();

    void on_btnClose_clicked();

private:
    Ui::MainWindow *ui;
    QSqlDatabase db;
    QSqlTableModel *modelT;
    QFileSystemModel *modelF;
    QMessageBox *msgBox;

    //Имя файла
        QString fileName;
    //Путь к файлу
        QString filePath;
};
#endif // MAINWINDOW_H
