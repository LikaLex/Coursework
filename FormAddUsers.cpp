//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop
#include "UsersForm.h"
#include "UnitQuestion.h"
#include "FormStart.h"
#include "FormCharacteristics.h"
#include "FormTest.h"
#include "FormAddUsers.h"
#include "UnitTest.h"
#include "UnitUser.h"
#include <ComObj.hpp>
#include <utilcls.h>

//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma resource "*.dfm"
TFAddUser *FAddUser;

extern AnsiString username;
//AnsiString CURRENT_DIRECTORY;
//username = ComboBoxUsers->Text;


//---------------------------------------------------------------------------
__fastcall TFAddUser::TFAddUser(TComponent* Owner)
	: TForm(Owner)
{
}
 extern AnsiString CURRENT_DIRECTORY;
 extern bool wasStart;


void addToCell(Variant Sheet,int row,int col,AnsiString value){
	Variant Cell;
	Cell=Sheet.OlePropertyGet("Cells").OlePropertyGet("Item",row,col);
	Cell.OlePropertySet("Value",StringToOleStr(value));
}

bool checkName = true;
bool checkSurname = true;
bool checkGroup = true;

//---------------------------------------------------------------------------
void __fastcall TFAddUser::AddButtonClick(TObject *Sender)
{

	/*wchar_t buffer[200];
	GetCurrentDirectory(sizeof(buffer),buffer);
	CURRENT_DIRECTORY=(AnsiString)buffer;*/
	Variant ExcelApplication,ExcelBooks,Sheet,Cell;
	int rowsCount;

	try{
		ExcelApplication=CreateOleObject("Excel.Application");
		ExcelBooks=ExcelApplication.OlePropertyGet("Workbooks").OlePropertyGet("Open",StringToOleStr(CURRENT_DIRECTORY+"\\Users.xlsx"));
		Sheet=ExcelBooks.OlePropertyGet("Worksheets",1);
		rowsCount=Sheet.OlePropertyGet("UsedRange").OlePropertyGet("Rows").OlePropertyGet("Count");
	}
	catch(...){
		Application->Title="Ошибка";
		ShowMessage("Ошибка при открытии файла\n"+CURRENT_DIRECTORY+"\\Users.xlsx"+"\nПроверьте наличие файла \"Users.xlsx\" в директории\n"+CURRENT_DIRECTORY);
		ExcelApplication.OleProcedure("Quit");
	}

	try{

		AnsiString Name=NameBox->Text;
		AnsiString Surname=SurnameBox->Text;
		AnsiString Group=GroupBox->Text;


		if (NameBox->Text == "" || NameBox->Text=="Имя") {
		 checkName = false;

		}
		else checkName = true;
		if (SurnameBox->Text == "" || SurnameBox->Text =="Фамилия") {
		 checkSurname = false;
		}
		else checkSurname = true;
		if (GroupBox->Text == "") {
		 checkGroup = false;
		}
		else checkGroup = true;

		if (checkName == true && checkSurname == true && checkGroup == true ) {



		FStart->ComboBoxUsers->Items->Add(Name);
		//FStart->ComboBoxUsers->Items->Add(Surname);
		//FStart->ComboBoxUsers->Items->Add(Group);

		addToCell(Sheet,rowsCount+1,1,Name);
		addToCell(Sheet,rowsCount+1,5,Surname);
		addToCell(Sheet,rowsCount+1,6,Group);
		addToCell(Sheet,rowsCount+1,2,0);
		addToCell(Sheet,rowsCount+1,3,0);
		addToCell(Sheet,rowsCount+1,4,0);
		addToCell(Sheet,rowsCount+1,7,0);
		addToCell(Sheet,rowsCount+1,8,0);

		addToCell(Sheet,rowsCount+1,11,0);
		addToCell(Sheet,rowsCount+1,12,0);
		addToCell(Sheet,rowsCount+1,13,0);
		addToCell(Sheet,rowsCount+1,14,0);
		addToCell(Sheet,rowsCount+1,15,0);
		addToCell(Sheet,rowsCount+1,16,0);
		addToCell(Sheet,rowsCount+1,17,0);
		addToCell(Sheet,rowsCount+1,18,0);
		addToCell(Sheet,rowsCount+1,19,0);
		addToCell(Sheet,rowsCount+1,20,0);

		addToCell(Sheet,rowsCount+1,21,0);
		addToCell(Sheet,rowsCount+1,22,0);
		addToCell(Sheet,rowsCount+1,23,0);
		addToCell(Sheet,rowsCount+1,24,0);
		addToCell(Sheet,rowsCount+1,25,0);
		addToCell(Sheet,rowsCount+1,26,0);
		addToCell(Sheet,rowsCount+1,27,0);
		addToCell(Sheet,rowsCount+1,28,0);
		addToCell(Sheet,rowsCount+1,29,0);
		addToCell(Sheet,rowsCount+1,30,0);

		addToCell(Sheet,rowsCount+1,31,0);
		addToCell(Sheet,rowsCount+1,32,0);
		addToCell(Sheet,rowsCount+1,33,0);
		addToCell(Sheet,rowsCount+1,34,0);
		addToCell(Sheet,rowsCount+1,35,0);
		addToCell(Sheet,rowsCount+1,36,0);
		addToCell(Sheet,rowsCount+1,37,0);
		addToCell(Sheet,rowsCount+1,38,0);
		addToCell(Sheet,rowsCount+1,39,0);
		addToCell(Sheet,rowsCount+1,40,0);

		addToCell(Sheet,rowsCount+1,41,0);
		addToCell(Sheet,rowsCount+1,42,0);
		addToCell(Sheet,rowsCount+1,43,0);
		addToCell(Sheet,rowsCount+1,44,0);
		addToCell(Sheet,rowsCount+1,45,0);
		addToCell(Sheet,rowsCount+1,46,0);
		addToCell(Sheet,rowsCount+1,47,0);
		addToCell(Sheet,rowsCount+1,48,0);
		addToCell(Sheet,rowsCount+1,49,0);
		addToCell(Sheet,rowsCount+1,50,0);

		addToCell(Sheet,rowsCount+1,51,0);
		addToCell(Sheet,rowsCount+1,52,0);
		addToCell(Sheet,rowsCount+1,53,0);
		addToCell(Sheet,rowsCount+1,54,0);
		addToCell(Sheet,rowsCount+1,55,0);
		addToCell(Sheet,rowsCount+1,56,0);
		addToCell(Sheet,rowsCount+1,57,0);
		addToCell(Sheet,rowsCount+1,58,0);
		addToCell(Sheet,rowsCount+1,59,0);
		addToCell(Sheet,rowsCount+1,60,0);

		Application->Title = "Добавление пользователя";
		ShowMessage("Пользователь "+NameBox->Text+" успешно добавлен");
		int count = FStart->ComboBoxUsers->Items->Count;
		   for (int i=0; i < count; i++) {
				if (FAddUser->NameBox->Text == FStart->ComboBoxUsers->Items->Strings[i]) {
					FStart->ComboBoxUsers->ItemIndex  =i;
					FStart->ComboBoxUsers->OnSelect(Sender);
					if (wasStart == false) {
						FAddUser->Hide();
					}
					break;
				}
		   }
		   ReturnButton->Click();
	   }
	   else{
			Application->Title="Ошибка при добавлении пользователя";
			if (checkName == false) ShowMessage("Введите имя!");
			if (checkSurname == false) ShowMessage("Введите фамилию!");
			if (checkGroup == false) ShowMessage("Введите номер группы!");

			ExcelApplication.OlePropertyGet("Workbooks").OlePropertyGet("Item",1).OleProcedure("Save");
			ExcelApplication.OleProcedure("Quit");
			return;
       }
		//NameBox->Visible = false;
		//OKButton->Visible = false;
		ExcelApplication.OlePropertyGet("Workbooks").OlePropertyGet("Item",1).OleProcedure("Save");
	}
	catch(...){
		Application->Title="Ошибка";
		ShowMessage("Ошибка при записи данных в файл\n"+CURRENT_DIRECTORY+"\\Users.xlsx");
	}

	ExcelApplication.OleProcedure("Quit");



	/*AnsiString WayToFile="d:\\курсовой проект\\Пользователи.xlsx";
	Variant ExcelApplication,ExcelBooks,Sheet,Cell;

	ExcelApplication=CreateOleObject("Excel.Application");
	ExcelBooks=ExcelApplication.OlePropertyGet("Workbooks").OlePropertyGet("Open",StringToOleStr(WayToFile));
	Sheet=ExcelBooks.OlePropertyGet("Worksheets",1);
	int rowsCount=Sheet.OlePropertyGet("UsedRange").OlePropertyGet("Rows").OlePropertyGet("Count");

	AnsiString Name=NameBox->Text;
	AnsiString Surname=SurnameBox->Text;
	AnsiString Group=GroupBox->Text;
	FStart->ComboBoxUsers->Items->Add(Name);
	FStart->ComboBoxUsers->Items->Add(Surname);
	FStart->ComboBoxUsers->Items->Add(Group);

	addToCell(Sheet,rowsCount+1,1,Name);
	addToCell(Sheet,rowsCount+1,5,Surname);
	addToCell(Sheet,rowsCount+1,6,Group);


	ExcelApplication.OlePropertyGet("Workbooks").OlePropertyGet("Item",1).OleProcedure("Save");

	ExcelApplication.OleProcedure("Quit");*/


	NameBox->Clear();
	SurnameBox->Clear();
	GroupBox->Clear();
}

//---------------------------------------------------------------------------
void __fastcall TFAddUser::NameBoxChange(TObject *Sender)
{
//NameBox->Clear();
}
//---------------------------------------------------------------------------
void __fastcall TFAddUser::SurnameBoxChange(TObject *Sender)
{
//SurnameBox->Clear();
}
//---------------------------------------------------------------------------
void __fastcall TFAddUser::GroupBoxChange(TObject *Sender)
{
//GroupBox->Clear();
}
//---------------------------------------------------------------------------
void __fastcall TFAddUser::ReturnButtonClick(TObject *Sender)
{




	try{

	FAddUser->Visible=false;
	FStart->Visible=true;
	FStart->Show();
	}
	catch(...){}

}
//---------------------------------------------------------------------------
void __fastcall TFAddUser::FormClose(TObject *Sender, TCloseAction &Action)
{

	try{

	//FAddUser->Visible=false;
	FStart->Visible=true;
	FStart->Show();
	}
	catch(...){}
}
//---------------------------------------------------------------------------
void __fastcall TFAddUser::NameBoxClick(TObject *Sender)
{
	if (NameBox->Text == "Имя") NameBox->Text="";
}
//---------------------------------------------------------------------------
void __fastcall TFAddUser::SurnameBoxClick(TObject *Sender)
{
	if (SurnameBox->Text=="Фамилия") SurnameBox->Text="";
}
//---------------------------------------------------------------------------
void __fastcall TFAddUser::NameBoxExit(TObject *Sender)
{
	if (NameBox->Text=="") NameBox->Text="Имя";
}
//---------------------------------------------------------------------------
void __fastcall TFAddUser::SurnameBoxExit(TObject *Sender)
{
	if (SurnameBox->Text=="") SurnameBox->Text="Фамилия";
}
//---------------------------------------------------------------------------

void __fastcall TFAddUser::FormActivate(TObject *Sender)
{
	WelkomeImage->Left = (FAddUser->ClientWidth - WelkomeImage->Width)/2;
	NameLabel->Left = WelkomeImage->Left;
	SurnameLabel->Left = WelkomeImage->Left;
	GroupLabel->Left= WelkomeImage->Left;

	NameBox->Left = WelkomeImage->Left + WelkomeImage->Width - NameBox->Width;
	SurnameBox->Left = WelkomeImage->Left + WelkomeImage->Width - SurnameBox->Width;
	GroupBox->Left = WelkomeImage->Left + WelkomeImage->Width - GroupBox->Width;
}
//---------------------------------------------------------------------------

