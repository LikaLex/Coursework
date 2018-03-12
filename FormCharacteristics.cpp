//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop

#include "UnitQuestion.h"
#include "FormStart.h"
#include "FormCharacteristics.h"
#include "FormTest.h"
#include "UnitTest.h"
#include "FormReports.h"

#include <stdio.h>
//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma resource "*.dfm"
TFCharacteristics *FCharacteristics;
//---------------------------------------------------------------------------
__fastcall TFCharacteristics::TFCharacteristics(TComponent* Owner)
	: TForm(Owner)
{
}
//---------------------------------------------------------------------------
extern AnsiString CURRENT_DIRECTORY;
int maxQuestions;
int SECONDS_FOR_ANSWER;

void __fastcall TFCharacteristics::EditQuestionsAmountClick(TObject *Sender)
{
	EditQuestionsAmount->Text="";
}
//---------------------------------------------------------------------------
void __fastcall TFCharacteristics::EditQuestionsAmountEnter(TObject *Sender)
{
	EditQuestionsAmount->Text="";
}
//---------------------------------------------------------------------------
void __fastcall TFCharacteristics::EditQuestionsAmountExit(TObject *Sender)
{
	if (EditQuestionsAmount->Text == "") {
		EditQuestionsAmount->Text="3";
	}
}
//---------------------------------------------------------------------------
void __fastcall TFCharacteristics::EditVariantsAmountClick(TObject *Sender)
{
	EditVariantsAmount->Text="";
}
//---------------------------------------------------------------------------
void __fastcall TFCharacteristics::EditVariantsAmountEnter(TObject *Sender)
{
	EditVariantsAmount->Text="";
}
//---------------------------------------------------------------------------
void __fastcall TFCharacteristics::EditVariantsAmountExit(TObject *Sender)
{
	if (EditVariantsAmount->Text == "") {
		EditVariantsAmount->Text="4";
	}
}
//---------------------------------------------------------------------------
bool isRightData(AnsiString data,int maxValue){
	int i,length=data.Length();
	bool isRight=true;
	for (i = 1; i <= length; i++) {
		if (data[i] > 57 || data[i]<48) {
			isRight=false;
			break;
		}
	}
	if (isRight == true) {
		int value=data.ToInt();
		if (value <= maxValue &&value > 0) {
			return true;
		}
		else return false;
	}
	else return false;
}

void __fastcall TFCharacteristics::ButtonGoBackClick(TObject *Sender)
{
	if (ComboBoxTheme->Text == "Мультимедийные вопросы" && CheckBoxMedia->Checked==false ) {
		CheckBoxMedia->Checked = true;
	}

	try{
		const int MAX_VARIANTS_AMOUNT=10;

			if (isRightData(EditVariantsAmount->Text,MAX_VARIANTS_AMOUNT) == true) {
				saveSettings();
				FCharacteristics->Close();
				FStart->Show();
			}
			else {
				EditVariantsAmount->Text="";
				Application->Title="Ошибка";
				ShowMessage("Неверное значение");
			}
	}
	catch(...){

	}
}
//---------------------------------------------------------------------------
void saveSettings(){
	Settings settings;

	settings.questionsAmount=FCharacteristics->UpDown1->Position;
	settings.variantsAmount=FCharacteristics->UpDown2->Position;
	settings.themeIndex=FCharacteristics->ComboBoxTheme->ItemIndex;
	settings.isTimeLimited=FCharacteristics->CheckBoxTime->Checked;
	settings.timeIndex=FCharacteristics->RadioGroupTime->ItemIndex;
	settings.autoCreateDocument=FCharacteristics->CheckBoxAutoCreateDocument->Checked;
	settings.autoSaveTest=FCharacteristics->CheckBoxAutoSaveTest->Checked;
	settings.areImages=FCharacteristics->CheckBoxImages->Checked;
	settings.areMedia=FCharacteristics->CheckBoxMedia->Checked;

	AnsiString wayToFile=CURRENT_DIRECTORY+"\\resourses\\settings.txt";
	FILE* file=fopen(wayToFile.c_str(),"w");
	fwrite(&settings,sizeof(Settings),1,file);
	fclose(file);
}
void getSettings(){
	Settings settings;
	AnsiString wayToFile=CURRENT_DIRECTORY+"\\resourses\\settings.txt";
	try{
		FILE* file=fopen(wayToFile.c_str(),"r");
		if (file!=NULL){
			fseek(file,0,SEEK_SET);
			fread(&settings,sizeof(Settings),1,file);
		}
		else {
			settings.questionsAmount=5;
			settings.variantsAmount=6;
			settings.themeIndex=0;
			settings.isTimeLimited=false;
			settings.timeIndex=0;
			settings.autoCreateDocument=true;
			settings.autoSaveTest=true;
			settings.areImages=true;
			settings.areMedia=true;
        }
		fclose(file);

		FCharacteristics->EditQuestionsAmount->Text=IntToStr(settings.questionsAmount);
		FCharacteristics->EditVariantsAmount->Text=IntToStr(settings.variantsAmount);
		FCharacteristics->ComboBoxTheme->ItemIndex=settings.themeIndex;
		FCharacteristics->CheckBoxTime->Checked=settings.isTimeLimited;
		FCharacteristics->RadioGroupTime->ItemIndex=settings.timeIndex;
		FCharacteristics->CheckBoxAutoCreateDocument->Checked=settings.autoCreateDocument;
		FCharacteristics->CheckBoxAutoSaveTest->Checked=settings.autoSaveTest;
		FCharacteristics->CheckBoxImages->Checked=settings.areImages;
		FCharacteristics->CheckBoxMedia->Checked=settings.areMedia;
	}
	catch(...){
		FCharacteristics->EditQuestionsAmount->Text=IntToStr(5);
		FCharacteristics->EditVariantsAmount->Text=IntToStr(6);
		FCharacteristics->ComboBoxTheme->ItemIndex=0;
		FCharacteristics->CheckBoxTime->Checked=false;
		FCharacteristics->RadioGroupTime->ItemIndex=0;
		FCharacteristics->CheckBoxAutoCreateDocument->Checked=true;
		FCharacteristics->CheckBoxAutoSaveTest->Checked=true;
		FCharacteristics->CheckBoxMedia->Checked=true;
		FCharacteristics->CheckBoxImages->Checked=true;
	}
}
void __fastcall TFCharacteristics::FormActivate(TObject *Sender)
{
	if (ComboBoxTheme->Text=="Все") {
		UpDown1->Max=80;
	}
	else UpDown1->Max=15;
	getSettings();
	CheckBoxTime->OnClick(Sender);
}
//---------------------------------------------------------------------------
void __fastcall TFCharacteristics::FormHide(TObject *Sender)
{
	try{
		saveSettings();
	}
	catch(...){}
}
//---------------------------------------------------------------------------
void __fastcall TFCharacteristics::FormClose(TObject *Sender, TCloseAction &Action)
{
	if (FCharacteristics->ComboBoxTheme->Text == "Мультимедийные вопросы" && FCharacteristics->CheckBoxMedia->Checked==false ) {
		FCharacteristics->CheckBoxMedia->Checked = true;
	}
	try{
		saveSettings();
		FCharacteristics->Visible=false;
		FStart->Visible=true;
		FStart->Show();
	}
	catch(...){}
}
//---------------------------------------------------------------------------
void __fastcall TFCharacteristics::ComboBoxThemeChange(TObject *Sender)
{
	if (ComboBoxTheme->Text=="Все") {
		UpDown1->Max=80;
	}
	else UpDown1->Max=15;

	if (StrToInt(EditQuestionsAmount->Text) > UpDown1->Max) {
		EditQuestionsAmount->Text=IntToStr(UpDown1->Max);
	}
}
//---------------------------------------------------------------------------
void __fastcall TFCharacteristics::CheckBoxTimeClick(TObject *Sender)
{
	if (CheckBoxTime->Checked==true) {
		RadioGroupTime->Enabled=true;
		switch (RadioGroupTime->ItemIndex) {
			case 0: SECONDS_FOR_ANSWER=30; break;
			case 1: SECONDS_FOR_ANSWER=45; break;
			case 2: SECONDS_FOR_ANSWER=60; break;
			default: SECONDS_FOR_ANSWER=30;
		}
	}
	else RadioGroupTime->Enabled=false;
}
//---------------------------------------------------------------------------
void __fastcall TFCharacteristics::ButtonReportsClick(TObject *Sender)
{
	FReports->Show();
	FCharacteristics->Hide();
}
//---------------------------------------------------------------------------

