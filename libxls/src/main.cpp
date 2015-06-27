#include "XlsHelper.h"
#include <iostream>
#include <time.h>
void test_number()
{
	clock_t t1 = clock();
	XlsHelper xls("example.xls");
	for (int col = 0; col < 1000; col++)
	{
		for (int row = 0; row < 1000; row++)
		{
			xls.SetCellValue(row, col, (row+1)*(col+1));
		}
		printf("col %d\n", col);
	}
	clock_t t2 = clock();
	std::cout << "time: " << (double)(t2 - t1) / CLOCKS_PER_SEC << " sec\n" << std::endl;
	xls.Show();
	xls.Save();
	clock_t t3 = clock();
	std::cout << "time: " << (double)(t3 - t2) / CLOCKS_PER_SEC << " sec\n" << std::endl;
}

void test_string()
{
	clock_t t1 = clock();
	XlsHelper xls("example.xls");
	for (int col = 0; col < 1000; col++)
	{
		for (int row = 0; row < 1000; row++)
		{
			char value[128] = "";
			sprintf(value, "%d", (row+1)*(col+1));
			xls.SetCellValue(row, col, value);
		}
		printf("col %d\n", col);
	}

	clock_t t2 = clock();
	std::cout << "time: " << (double)(t2 - t1) / CLOCKS_PER_SEC << " sec\n" << std::endl;
	xls.Show();
	xls.Save();
	clock_t t3 = clock();
	std::cout << "time: " << (double)(t3 - t2) / CLOCKS_PER_SEC << " sec\n" << std::endl;
}
int main()
{
	//test_number();
	//test_string();

	XlsHelper xls("sample.xls");
	
	xls.Show();
}