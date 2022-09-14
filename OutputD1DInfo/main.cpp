#include <iostream>
#include <fstream>
#include <string>
#include <stdlib.h>
#include <math.h>
#include <unistd.h>
#include <vector>
#include <sstream>
#include "BasicExcel.hpp"
#include "GetFunctions.hpp"
#define MAX_NUM_OF_FILES (5000) //Max number of files this program can handle
#define sensitivity (28.5)

using namespace std;

int main(){

    char cwd[1000];
    getcwd(cwd, sizeof(cwd));             //Find current working directory

    vector<string> files;
    getFiles(cwd, files);                 //Get file names of all files in current folder
    printf("Remaining following files in the current directory: \n%s\n\n",cwd);

    string D1Dfiles[MAX_NUM_OF_FILES];    //Filter out all the .D1D files
    int numofD1D = 0;
	for(unsigned int i = 0; i < files.size(); i++){
		if(files[i].find(".d1d")!=std::string::npos||files[i].find(".D1D")!=std::string::npos){
			D1Dfiles[numofD1D] = (files[i]);
			numofD1D++;
		}
	}

	if(numofD1D == 0){                    //If there is no .d1d file, exit directly.
		cout<<"No d1d file found. Press any key to exit."<<endl;
		while(getchar())
			return -1;
	}

	for(int i = 0; i < numofD1D; i++)     //List out all the files.
		cout<<D1Dfiles[i]<<endl;
	cout<<endl<<endl;

	cout<<"Output information:"<<endl;

	 for(int i = 0; i < numofD1D; i++){

		 cout<<D1Dfiles[i]<<endl<<"Date = "<<getDate(D1Dfiles[i])<<endl<<"Xo = "<<getXo(D1Dfiles[i])<<endl<<"Yo = "<<getYo(D1Dfiles[i])<<endl<<"Range = "<<round(getWidth(D1Dfiles[i]))<<endl<<"Angle = "<<getAngle(D1Dfiles[i])<<endl<<"Energy = "<<round(getEnergy(D1Dfiles[i]))<<endl<<"Gate Time = "<<getGT(D1Dfiles[i])<<endl<<"EL = "<<getEL(D1Dfiles[i])<<endl<<"L2 = "<<getL2(D1Dfiles[i])<<endl<<"L13 = "<<getL13(D1Dfiles[i])<<endl<<"W = "<<getW(D1Dfiles[i])<<endl<<endl;
	 }
	 cout<<endl;


	YExcel::BasicExcel OutInfo;
	OutInfo.New(1);

	YExcel::BasicExcelWorksheet* sheet = OutInfo.GetWorksheet("Sheet1");

	//Write title
	sheet->Cell(0,0)->SetString("Type");  //Type
	sheet->Cell(0,1)->SetString("Filename");  //Filename
	sheet->Cell(0,2)->SetString("Date");  //Date
	sheet->Cell(0,3)->SetString("Xo");  //Xo
	sheet->Cell(0,4)->SetString("Yo");  //Yo
	sheet->Cell(0,5)->SetString("Range");  //Range
	sheet->Cell(0,6)->SetString("Angle");  //Angle
	sheet->Cell(0,7)->SetString("Energy");  //Energy
	sheet->Cell(0,8)->SetString("Rounded Energy");  //Rounded energy
	sheet->Cell(0,9)->SetString("Gate Time");  //Gate time
	sheet->Cell(0,10)->SetString("EL");  //Entrance Lens
	sheet->Cell(0,11)->SetString("L2");  //L2
	sheet->Cell(0,12)->SetString("L13");  //L13
	sheet->Cell(0,13)->SetString("W");  //Wehnelt
	//Write in data
	if (sheet){
		for(int i = 0; i < numofD1D; i++){
			sheet->Cell(i+1,0)->SetString("D1D");  //Type
			sheet->Cell(i+1,1)->SetString(D1Dfiles[i].c_str());  //Filename
			sheet->Cell(i+1,2)->SetString(getDate(D1Dfiles[i]).c_str());  //Date
			sheet->Cell(i+1,3)->SetDouble(getXo(D1Dfiles[i]));  //Xo
			sheet->Cell(i+1,4)->SetDouble(getYo(D1Dfiles[i]));  //Yo
			sheet->Cell(i+1,5)->SetDouble(round(getWidth(D1Dfiles[i])));  //Range
			sheet->Cell(i+1,6)->SetDouble(getAngle(D1Dfiles[i]));  //Angle
			sheet->Cell(i+1,7)->SetDouble(getEnergy(D1Dfiles[i]));  //Energy
			sheet->Cell(i+1,8)->SetDouble(round(getEnergy(D1Dfiles[i])));  //Rounded energy
			sheet->Cell(i+1,9)->SetDouble((getGT(D1Dfiles[i])));  //Gate time
			sheet->Cell(i+1,10)->SetDouble((getEL(D1Dfiles[i])));  //Entrance Lens
			sheet->Cell(i+1,11)->SetDouble((getL2(D1Dfiles[i])));  //L2
			sheet->Cell(i+1,12)->SetDouble((getL13(D1Dfiles[i])));  //L13
			sheet->Cell(i+1,13)->SetDouble((getW(D1Dfiles[i])));  //Wehnelt
		}
	}

	char XLSN[30];
	strcpy(XLSN,"D1DInfo (");
	string d1dFN = D1Dfiles[0];
	d1dFN = d1dFN.substr(0,8);
	strcat(XLSN,d1dFN.c_str()); //using c_str convert std::string to char*
	strcat(XLSN,").xls");
	OutInfo.SaveAs(XLSN);

	cout<<"Read completed. Press any key to exit."<<endl;
	while(getchar())
	   	return 0;

}
