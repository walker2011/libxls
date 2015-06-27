#include "XlsHelper.h"

#include <Windows.h>
XlsHelper::XlsHelper(std::string workbookname)
	:workbookname_(workbookname)
	,workbook_(NULL)
	,worksheet_(NULL)
{
	this->CreateWookbook();
	this->AddWorksheet("Sheet1");
	this->AddWorksheet("Sheet2");
	this->AddWorksheet("Sheet3");
}

XlsHelper::~XlsHelper()
{
	this->Save();

	if (workbook_)
	{
		workbook_->release();
	}
}

bool XlsHelper::CreateWookbook()
{
	workbook_ = xlCreateBook();
	if (workbook_)
	{
		return true;
	}
	else
	{
		return false;
	}
}

bool XlsHelper::AddWorksheet(std::string worksheetname)
{
	if (workbook_)
	{
		worksheet_ = NULL;
		worksheet_ = workbook_->addSheet(worksheetname.data());
	}

	return workbook_!=NULL;
}

bool XlsHelper::SetWorksheet(int index)
{
	if (workbook_ == NULL)
		return false;

	if (index < 0 
		|| index > workbook_->sheetCount())
		return false;

	worksheet_ = workbook_->getSheet(index);
	return true;
}

void XlsHelper::SetCellValue(int row, int col, std::string value)
{
	worksheet_->writeStr(row, col, value.data());
}

void XlsHelper::SetCellValue(int row, int col, double value)
{
	worksheet_->writeNum(row, col, value);
}

std::string XlsHelper::GetCellString(int row, int col)
{
	return std::string(worksheet_->readStr(row, col));
}

double XlsHelper::GetCellNumber(int row, int col)
{
	return worksheet_->readNum(row, col);
}

bool XlsHelper::Save()
{
	if (workbook_ == NULL)
		return false;

	return workbook_->save(workbookname_.data());
}

void XlsHelper::Show()
{
	::ShellExecute(NULL, "open", workbookname_.data(), NULL, NULL, SW_SHOW);  
}

std::string XlsHelper::GetErrorMessage()
{
	const char* msg =NULL;

	if (workbook_)
	{
		msg = workbook_->errorMessage();
	}	

	return std::string(msg);
}