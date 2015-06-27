#ifndef XLS_HELPER_H
#define XLS_HELPER_H

#include <string>
#include "libxl.h"
using namespace libxl;
class XlsHelper
{
public:
	XlsHelper(std::string workbookname);
	~XlsHelper();

	//workbook
	bool CreateWookbook();
	bool Save();
	void Show();

	//worksheet
	bool AddWorksheet(std::string worksheetname);
	bool SetWorksheet(int index);

	//cell
	void SetCellValue(int row, int col, std::string value);
	void SetCellValue(int row, int col, double value);
	std::string GetCellString(int row, int col);
	double GetCellNumber(int row, int col);
	
	//error message
	std::string GetErrorMessage();
private:
	Book* workbook_;
	Sheet* worksheet_;
	std::string workbookname_;
};
#endif