//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop

#include "FormReports.h"
#include "FormCharacteristics.h"
#include "FormStart.h"

#include <direct.h>
//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma resource "*.dfm"
TFReports *FReports;
//---------------------------------------------------------------------------
__fastcall TFReports::TFReports(TComponent* Owner)
	: TForm(Owner)
{
}
//---------------------------------------------------------------------------
extern AnsiString CURRENT_DIRECTORY;

void __fastcall TFReports::FormActivate(TObject *Sender)
{
	bool wasAdded;
	getFiles("\\Documents",FReports->TreeViewReports,".docx",&wasAdded);
}
//---------------------------------------------------------------------------
void __fastcall TFReports::FormClose(TObject *Sender, TCloseAction &Action)
{
	FCharacteristics->Visible=true;
	FReports->Visible=false;
}
//---------------------------------------------------------------------------
void __fastcall TFReports::ButtonShowClick(TObject *Sender)
{
	try{
		UnicodeString FileName=CURRENT_DIRECTORY+"\\Documents\\"+TreeViewReports->Selected->Text+".docx";
		ShellExecute(NULL,L"open",FileName.w_str(),NULL,NULL,SW_NORMAL);
	}
	catch(...){
		Application->Title="Ошибка";
		ShowMessage("Выберите отчет!");
	}
}
//---------------------------------------------------------------------------
void __fastcall TFReports::ButtonDeleteClick(TObject *Sender)
{
	try{
		AnsiString FileName=CURRENT_DIRECTORY+"\\Documents\\"+TreeViewReports->Selected->Text+".docx";
		DeleteFile(FileName.c_str());
		FReports->Activate();
	}
	catch(...){
		Application->Title="Ошибка";
		ShowMessage("Выберите отчет!");
	}
}
//---------------------------------------------------------------------------
void __fastcall TFReports::ButtonDeleteAllClick(TObject *Sender)
{
	try{
		TTreeNode* node=TreeViewReports->Items->GetFirstNode();
		while (node!=NULL){
			AnsiString directory=CURRENT_DIRECTORY+"\\Documents\\"+node->Text+".docx";
			DeleteFile(directory.c_str());
			node=node->GetNext();
		}
		FReports->Activate();
	}
	catch(...){
		Application->Title="Ошибка";
		ShowMessage("Ошибка при удалении");
	}
}
//---------------------------------------------------------------------------
void __fastcall TFReports::ButtonBackClick(TObject *Sender)
{
	FReports->Close();
}
//---------------------------------------------------------------------------

