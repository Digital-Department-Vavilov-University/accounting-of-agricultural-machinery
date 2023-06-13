#include "MainForm.h"
#include <time.h>

#define CSV_DELIMETER ";"
#define CSV_DELIMETER_CHR ';'

using namespace System::Text;

System::Void Project1::MainForm::MainForm_Load(System::Object^ sender, System::EventArgs^ e)
{
    this->sumBelowCells = gcnew System::Collections::Generic::List<SumBelow^>();
    this->rowCalculationData = gcnew System::Collections::Generic::Dictionary<int, RowData^>();

    CreateGridData();

    return System::Void();
}

void Project1::MainForm::CreateGridData()
{
    LoadDataFromTemplate(this->grid);
}

System::Void Project1::MainForm::saveToolStripMenuItem_Click(System::Object^ sender, System::EventArgs^ e)
{
    SaveFileDialog^ saveFileDialog = gcnew SaveFileDialog();
    saveFileDialog->Filter = "(csv file(*.csv)| *.csv";
    try
    {
        if (saveFileDialog->ShowDialog() == System::Windows::Forms::DialogResult::OK)
        {
            String^ path = saveFileDialog->FileName;

            String^ collectedData = GetCSVData(this->grid);

            System::IO::File::WriteAllText(path, collectedData, System::Text::Encoding::UTF8);

            MessageBox::Show("Таблица успешно сохранена по пути\n" + path);
        }
        else {
            return System::Void();
        }
    }
    catch (Exception^ ex) {
        MessageBox::Show(this, "Не удалось сохранить. Возможно файл открыт в другой программе.", "Ошибка", MessageBoxButtons::OK, MessageBoxIcon::Error);
    }

    return System::Void();
}

System::String^ Project1::MainForm::GetCSVData(DataGridView^ grid)
{
    System::Text::StringBuilder^ sb = gcnew System::Text::StringBuilder();

    sb->Append(grid->TopLeftHeaderCell->Value->ToString());
    sb->Append(CSV_DELIMETER);
    for (int i = 0; i < grid->Columns->Count; i++) {
        System::Object^ header = grid->Columns[i]->HeaderCell->Value;
        if (header == nullptr) header = "";
        sb->Append(header->ToString());
        sb->Append(CSV_DELIMETER);
    }
    sb->Append(CSV_DELIMETER);
    sb->Append("\n");

    for (int i = 0; i < grid->Rows->Count; i++) {
        System::Object^ header = grid->Rows[i]->HeaderCell->Value;
        if (header == nullptr) header = "";
        sb->Append(header->ToString());
        sb->Append(CSV_DELIMETER);

        DataGridViewRow^ row = grid->Rows[i];
        for (int j = 0; j < row->Cells->Count; j++) {
            System::Object^ val = row->Cells[j]->Value;
            if (val == nullptr) val = "";
            sb->Append(row->Cells[j]->Value);
            sb->Append(CSV_DELIMETER);
        }
        sb->Append("\n");
    }

    return sb->ToString();
}

System::Void Project1::MainForm::grid_CellEndEdit(System::Object^ sender, System::Windows::Forms::DataGridViewCellEventArgs^ e)
{
    DataGridViewCell^ cell = this->grid->CurrentCell;
    String^ text = cell->Value->ToString();

    if (!ValidateCellText(text, cell)) return;

    RecalculateData();

    return System::Void();
}

bool Project1::MainForm::ValidateCellText(String^ text, DataGridViewCell^ cell)
{
    if (text->Length == 0) {
        cell->Value = 0;
        MessageBox::Show("Введите натуральное число");
        return false;
    }

    for (int i = 0; i < text->Length; i++) {
        if (!Char::IsDigit(text[i])) {
            cell->Value = 0;
            MessageBox::Show("Введите натуральное число");
            return false;
        }
    }

    if (text->Length > 15) {
        cell->Value = text->Substring(0, 15);
        MessageBox::Show("Слишком большое значение ячейки");
        return false;
    }

    return true;
}

void Project1::MainForm::RecalculateData()
{
    for each (RowData^ data in rowCalculationData->Values) {
        int beginData = Int64::Parse(data->beginDataCell->Value->ToString());
        int receivedData = Int64::Parse(data->receivedDataCell->Value->ToString());
        int goneData = Int64::Parse(data->goneDataCell->Value->ToString());

        int result = beginData + receivedData - goneData;

        data->resultCell->Value = result;
    }

    for (int i = 0; i < sumBelowCells->Count; i++) {
        DataGridViewCell^ cell = sumBelowCells[i]->cell;
        int rowsAmount = sumBelowCells[i]->belowCellsCount;
        int column = cell->ColumnIndex;
        int row = cell->RowIndex;

        long long sum = 0;
        for (int j = 0; j < rowsAmount; j++) {
            sum += Int64::Parse(this->grid->Rows[row + 1 + j]->Cells[column]->Value->ToString());
        }

        cell->Value = sum;
    }
}

void Project1::MainForm::LoadDataFromTemplate(System::Windows::Forms::DataGridView^ grid)
{
    array<String^>^ rows = System::IO::File::ReadAllLines("template.csv", System::Text::Encoding::UTF8);
    array<String^>^ row0 = rows[0]->Split(CSV_DELIMETER_CHR);
    grid->TopLeftHeaderCell->Value = row0[0];
    for (int i = 1; i < row0->Length; i++) {
        grid->Columns->Add(row0[i], row0[i]);
        grid->Columns[i - 1]->SortMode = DataGridViewColumnSortMode::NotSortable;
    }
    grid->Rows->Add(rows->Length - 1);
    for (int i = 1; i < rows->Length; i++) {
        array<String^>^ cells = rows[i]->Split(CSV_DELIMETER_CHR);

        ProcessCell(this->grid->Rows[i - 1]->HeaderCell, 0, 0, cells[0]);

        for (int j = 1; j < cells->Length; j++) {
            ProcessCell(this->grid->Rows[i - 1]->Cells[j - 1], i, j, cells[j]);
        }
    }
}

void Project1::MainForm::ProcessCell(System::Windows::Forms::DataGridViewCell^ cell, int row, int column, String^ rawCellText)
{
    RegularExpressions::Regex^ tagRegex = gcnew RegularExpressions::Regex("<\.*?>");

    String^ textData = tagRegex->Replace(rawCellText, String::Empty);
    cell->Value = textData;
    RegularExpressions::MatchCollection^ matches = tagRegex->Matches(rawCellText);    
    for (int i = 0; i < matches->Count; i++) {
        ProcessTag(cell, matches[i]->Value);
    }

    if (!rawCellText->Contains("<EDIT>") && row > 0 && column > 0)
    {
        cell->ReadOnly = true;
    }
}

void Project1::MainForm::ProcessTag(System::Windows::Forms::DataGridViewCell^ cell, System::String^ tag) {
    if (tag->Equals("<BOLD>")) {
        cell->Style->Font = gcnew System::Drawing::Font(grid->Font, System::Drawing::FontStyle::Bold);
    }
    else if (tag->Equals("<BACK YELLOW>")) {
        cell->Style->BackColor = System::Drawing::Color::LightYellow;
    }
    else if (tag->Equals("<BACK BLUE>")) {
        cell->Style->BackColor = System::Drawing::Color::LightBlue;
    }
    else if (tag->Equals("<BACK GREEN>")) {
        cell->Style->BackColor = System::Drawing::Color::LightGreen;
    }
    else if (tag->Equals("<BEGINDATA>")) {
        int row = cell->RowIndex;
        if (!this->rowCalculationData->ContainsKey(row)) this->rowCalculationData[row] = gcnew RowData();
        this->rowCalculationData[row]->beginDataCell = cell;
    }
    else if (tag->Equals("<RECEIVEDDATA>")) {
        int row = cell->RowIndex;
        if (!this->rowCalculationData->ContainsKey(row)) this->rowCalculationData[row] = gcnew RowData();
        this->rowCalculationData[row]->receivedDataCell = cell;
    }
    else if (tag->Equals("<GONEDATA>")) {
        int row = cell->RowIndex;
        if (!this->rowCalculationData->ContainsKey(row)) this->rowCalculationData[row] = gcnew RowData();
        this->rowCalculationData[row]->goneDataCell = cell;
    }
    else if (tag->Equals("<ENDDATA>")) {
        int row = cell->RowIndex;
        if (!this->rowCalculationData->ContainsKey(row)) this->rowCalculationData[row] = gcnew RowData();
        this->rowCalculationData[row]->resultCell = cell;
    }
    else if (tag->Contains("SUMBELOW")) {
        int rowsCount = Int32::Parse(tag->Substring(1, tag->Length - 2)->Split(' ')[1]);
        this->sumBelowCells->Add(gcnew SumBelow(cell, rowsCount));
    }
}