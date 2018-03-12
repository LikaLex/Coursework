//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop

#include "UsersForm.h"
#include "UnitQuestion.h"
#include "FormStart.h"
#include "FormCharacteristics.h"
#include "FormTest.h"
#include "FormAddUsers.h"
#include "RecordsForm.h"
#include "RatingForm.h"
#include "UnitTest.h"
#include "UnitUser.h"
#include <ComObj.hpp>
#include <utilcls.h>

//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma resource "*.dfm"
TUsersF *UsersF;
//AnsiString CURRENT_DIRECTORY;
 extern AnsiString CURRENT_DIRECTORY;
//---------------------------------------------------------------------------
__fastcall TUsersF::TUsersF(TComponent* Owner)
	: TForm(Owner)
{
}
	User user;
//---------------------------------------------------------------------------
void __fastcall TUsersF::LastResultButtonClick(TObject *Sender)
{
Application->Title="����������";
	ShowMessage("��������� ���������� ������������: " + IntToStr(user.getLastMark()));


}
//---------------------------------------------------------------------------

void __fastcall TUsersF::BackToTestButtonClick(TObject *Sender)
{


	try{

	UsersF->Visible=false;
	FStart->Visible=true;
	FStart->Show();
	}
	catch(...){}

}
//---------------------------------------------------------------------------

void __fastcall TUsersF::AllTestsButtonClick(TObject *Sender)
{
    Application->Title="����������";
	ShowMessage("���������� ���������� ������: " + IntToStr(user.getAllTestsResult()));


}
//---------------------------------------------------------------------------

void __fastcall TUsersF::AverageScoreButtonClick(TObject *Sender)
{
	Application->Title="����������";
	 ShowMessage("������� ���� ������������: " + FloatToStr(user.getAverageScore()));

}
//---------------------------------------------------------------------------

void __fastcall TUsersF::InformationButtonClick(TObject *Sender)
{
	Application->Title="����������";
		ShowMessage("���: " + user.getName() + "\n" + "�������: " + 	user.getSurname() + "\n" + "������: " + user.getGroup() + "\n" + "��������� ���������� ������������: " + IntToStr(user.getLastMark()) + "\n" + "���������� ���������� ������: " + IntToStr(user.getAllTestsResult()) + "\n" + "������� ���� ������������: " + FloatToStr(user.getAverageScore())+ "\n" + "�����, ����������� �� ����������� ���������� ������������: " + IntToStr(user.getLastTime())+" ������" + "\n" + "�����, ����������� ������������� �� ����������� ���� ������: " + IntToStr(user.getAllTime())+ " ������");

}
//---------------------------------------------------------------------------

void __fastcall TUsersF::FormActivate(TObject *Sender)
{   /*
	int amount = FStart->ComboBoxUsers->Items->Count;
	bool wasFounded=false;
	for (int i=0; i < amount; i++) {
	 AnsiString str = FStart->ComboBoxUsers->Items->Strings[i];
	 if (FStart->ComboBoxUsers->Text == str) {
	   wasFounded = true;
	   break;
	 }
	}
	if (wasFounded == false) {
		FAddUser->MemoHelp->Lines->Clear();
		FAddUser->MemoHelp->Lines->Add("��� �� ������� � ������ �������������");
		 FAddUser->MemoHelp->Lines->Add("����������, �������� �����������.");
		 FAddUser->WelkomeImage->Visible=false;
		 FAddUser->MemoHelp->Visible = true;
		 FAddUser->MemoHelp->Left = FAddUser->WelkomeImage->Left;
		FAddUser->Show();
		return;
	}   */

	FStart->Visible = false;
	AnsiString temp;
	AnsiString WayToFile;
	Variant ExcelApplication,ExcelBooks,Sheet,Cell;
	wchar_t buffer[200];
	GetCurrentDirectory(sizeof(buffer),buffer);
	CURRENT_DIRECTORY=(AnsiString)buffer;
	//WayToFile="d:\\�������� ������\\������������.xlsx";
	int i;
	int position = FStart->ComboBoxUsers->ItemIndex+1;
	//const int COLS_AMOUNT = 6;
	//Variant ExcelApplication,ExcelBooks,Sheet,Cell;
	int rowsCount;

	try{
		ExcelApplication=CreateOleObject("Excel.Application");
		ExcelBooks=ExcelApplication.OlePropertyGet("Workbooks").OlePropertyGet("Open",StringToOleStr(CURRENT_DIRECTORY+"\\Users.xlsx"));
		Sheet=ExcelBooks.OlePropertyGet("Worksheets",1);
		rowsCount=Sheet.OlePropertyGet("UsedRange").OlePropertyGet("Rows").OlePropertyGet("Count");
	}
	catch(...){
		Application->Title="������";
		ShowMessage("������ ��� �������� �����\n"+CURRENT_DIRECTORY+"\\Users.xlsx"+"\n��������� ������� ����� \"Users.xlsx\" � ����������\n"+CURRENT_DIRECTORY);
		ExcelApplication.OleProcedure("Quit");
	}

	try{
	temp=Sheet.OlePropertyGet("Cells").OlePropertyGet("Item",position,1);//����� ������ �((��� ������������)
	user.setName(temp);

	temp=Sheet.OlePropertyGet("Cells").OlePropertyGet("Item",position,2);//����� ������ � (��������� ������)
	user.setLastMark(StrToInt(temp));

	temp=Sheet.OlePropertyGet("Cells").OlePropertyGet("Item",position,3);//����� ������ C (���������� ���������� ������)
	user.setAllTestsResult(StrToInt(temp));

	temp=Sheet.OlePropertyGet("Cells").OlePropertyGet("Item",position,4);//����� ������ D (������� ����)
	user.setAverageScore(StrToFloat(temp));

	temp=Sheet.OlePropertyGet("Cells").OlePropertyGet("Item",position,5);//����� ������ E (������� ������������)
	user.setSurname(temp);

	temp=Sheet.OlePropertyGet("Cells").OlePropertyGet("Item",position,6);//����� ������ F (������ ������������)
	user.setGroup(temp);

	temp=Sheet.OlePropertyGet("Cells").OlePropertyGet("Item",position,7);//����� ������ F (������ ������������)
	user.setLastTime(StrToInt(temp));

	temp=Sheet.OlePropertyGet("Cells").OlePropertyGet("Item",position,8);//����� ������ F (������ ������������)
	user.setAllTime(StrToInt(temp)) ;
	}
	catch(...){
		Application->Title="������";
		ShowMessage(" ������ ������������ ���.���� ������ ������ ������������, ����� ������ �����������");
			UsersF->Hide();

		FStart->ComboBoxUsers->ItemIndex = 0;
		FAddUser->Show();
	}




	ExcelApplication.OleProcedure("Quit");
}
//---------------------------------------------------------------------------

void __fastcall TUsersF::FormClose(TObject *Sender, TCloseAction &Action)
{
	FAddUser->MemoHelp->Visible=false;

	try{

	UsersF->Visible=false;
	FStart->Visible=true;
	FStart->Show();
	}
	catch(...){}

}
//---------------------------------------------------------------------------

void __fastcall TUsersF::LastTimeButtonClick(TObject *Sender)
{
	Application->Title="����������";
	ShowMessage("�����, ����������� �� ����������� ���������� ������������: " + IntToStr(user.getLastTime())+" ������");
}
//---------------------------------------------------------------------------

void __fastcall TUsersF::AllTimeButtonClick(TObject *Sender)
{
	Application->Title="����������";
	ShowMessage("�����, ����������� ������������� �� ����������� ���� ������: " + IntToStr(user.getAllTime())+" ������");
}
//---------------------------------------------------------------------------


