#include "mainwindow.h"
#include "ui_mainwindow.h"
#include "QAxObject"
#include "QFileDialog"
#include "QStandardPaths"
#include "QDebug"

double EPS = 0.0001;
QString structParams [6] = {"СПИСОК ИСПОЛНИТЕЛЕЙ", "РЕФЕРАТ", "СОДЕРЖАНИЕ", "ВВЕДЕНИЕ", "ЗАКЛЮЧЕНИЕ", "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ"};

MainWindow::MainWindow(QWidget *parent)
    : QMainWindow(parent)
    , ui(new Ui::MainWindow)
{
    ui->setupUi(this);
}

MainWindow::~MainWindow()
{
    delete ui;
}


void MainWindow::on_pushButton_clicked()
{
    ui->textBrowser->setText("");
    QString result = "РЕЗУЛЬТАТЫ АНАЛИЗА\n";
    QString text = "";
    bool fontName = 1;
    bool fontSize = 1;
    bool fontColor = 1;
    bool margins[4] = {1, 1, 1, 1};
    QString filename = QFileDialog::getOpenFileName(this, "Get Word", QStandardPaths::writableLocation(QStandardPaths::HomeLocation), tr("Word (*.docx)"));
    QAxObject *wordApplication = new QAxObject ("Word.Application");
    QAxObject *wordDoc = wordApplication->querySubObject("Documents()");
    QAxObject *doc = wordDoc->querySubObject("Add(QVariant)", filename);
    QAxObject *word, *font;
    QAxObject *words = doc->querySubObject("Words");
    int wordCount = words->property("Count").toInt();
    for (int a = 1; a <= wordCount; a++) {
        word = words->querySubObject("Item(int)", a)->querySubObject("FormattedText");
        font = word->querySubObject("Font");
        if(font->property("Name").toString() != ("Times New Roman")) fontName = 0;
        if(font->property("Size").toInt() < 12) fontSize = 0;
        if(font->property("Color").toInt() != -587137025) fontColor = 0;
        text += word->property("Text").toString();
    }

    result += "Путь к файлу: " + filename + "\n";

    //1 Структурные элементы и регистр стрктурных элементов
    if(!text.contains(structParams[0], Qt::CaseInsensitive)) result += "Структурный элемент " + structParams[0] + " не найден\n";
    else if(!text.contains(structParams[0], Qt::CaseSensitive)) result += "Название структурного элемента " + structParams[0] + " должно быть написано заглавными буквами\n";
    if(!text.contains(structParams[1], Qt::CaseInsensitive)) result += "Структурный элемент " + structParams[1] + " не найден\n";
    else if(!text.contains(structParams[1], Qt::CaseSensitive)) result += "Название структурного элемента " + structParams[1] + " должно быть написано заглавными буквами\n";
    if(!text.contains(structParams[2], Qt::CaseInsensitive)) result += "Структурный элемент " + structParams[2] + " не найден\n";
    else if(!text.contains(structParams[2], Qt::CaseSensitive)) result += "Название структурного элемента " + structParams[2] + " должно быть написано заглавными буквами\n";
    if(!text.contains(structParams[3], Qt::CaseInsensitive)) result += "Структурный элемент " + structParams[3] + " не найден\n";
    else if(!text.contains(structParams[3], Qt::CaseSensitive)) result += "Название структурного элемента " + structParams[3] + " должно быть написано заглавными буквами\n";
    if(!text.contains(structParams[4], Qt::CaseInsensitive)) result += "Структурный элемент " + structParams[4] + " не найден\n";
    else if(!text.contains(structParams[4], Qt::CaseSensitive)) result += "Название структурного элемента " + structParams[4] + " должно быть написано заглавными буквами\n";
    if(!text.contains(structParams[5], Qt::CaseInsensitive)) result += "Структурный элемент " + structParams[5] + " не найден\n";
    else if(!text.contains(structParams[5], Qt::CaseSensitive)) result += "Название структурного элемента " + structParams[5] + " должно быть написано заглавными буквами\n";
    if(text.contains(structParams[0], Qt::CaseInsensitive) && text.contains(structParams[1], Qt::CaseInsensitive) && text.contains(structParams[2], Qt::CaseInsensitive)
            && text.contains(structParams[3], Qt::CaseInsensitive) && text.contains(structParams[4], Qt::CaseInsensitive) && text.contains(structParams[5], Qt::CaseInsensitive))
        result += "Все основные структурные элементы найдены\n";
    if(text.contains(structParams[0], Qt::CaseSensitive) && text.contains(structParams[1], Qt::CaseSensitive) && text.contains(structParams[2], Qt::CaseSensitive)
            && text.contains(structParams[3], Qt::CaseSensitive) && text.contains(structParams[4], Qt::CaseSensitive) && text.contains(structParams[5], Qt::CaseSensitive))
        result += "Все основные структурные элементы названы верно\n";

    //2 Шрифт
    if(fontName && fontSize && fontColor) result += "Все требования к шрифту соблюдены\n";
    else{
        if(!fontName) result += "Рекомендуемый тип шрифт Times New Roman\n";
        if(!fontSize) result += "Размер шрифта должен быть не меньше 12 пт\n";
        if(!fontColor) result += "Цвет шрифта должен быть чёрным\n";
    }

    //3 Поля
    if(doc->querySubObject("PageSetup")->property("BottomMargin").toDouble() - 56.7 > EPS) margins[0] = 0;
    if(doc->querySubObject("PageSetup")->property("TopMargin").toDouble() - 56.7 > EPS) margins[1] = 0;
    if(doc->querySubObject("PageSetup")->property("LeftMargin").toDouble() - 85.05 > EPS) margins[2] = 0;
    if(doc->querySubObject("PageSetup")->property("RightMargin").toDouble() - 42.5 > EPS) margins[3] = 0;

    if(margins[0] && margins[1] && margins[2] && margins[3]) result += "Размеры всех полей верны\n";
    else{
        if(!margins[2]) result += "Размер левого поля должен быть 30 мм\n";
        if(!margins[3]) result += "Размер правого поля должен быть 15 мм\n";
        if(!margins[1]) result += "Размер верхнего поля должен быть 20 мм\n";
        if(!margins[0]) result += "Размер нижнего поля должен быть 20 мм\n";
    }

    ui->textBrowser->setText(result);
    delete font;
    delete word;
    wordApplication->dynamicCall("Quit()");
    delete words;
    delete doc;
    delete wordDoc;
    delete wordApplication;
}

