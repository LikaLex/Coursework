//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop

#include "UnitQuestion.h"
#include "FormStart.h"
#include "FormCharacteristics.h"
#include "FormTest.h"
#include "FormAddUsers.h"
#include "RecordsForm.h"
#include "RatingForm.h"
#include "TestsRatingForm.h"
#include "UnitTest.h"
#include "UnitUser.h"

//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma resource "*.dfm"
TFormTestsRating *FormTestsRating;
extern AnsiString  CURRENT_DIRECTORY;
//---------------------------------------------------------------------------
__fastcall TFormTestsRating::TFormTestsRating(TComponent* Owner)
	: TForm(Owner)
{
}
//---------------------------------------------------------------------------
void __fastcall TFormTestsRating::BackButtonClick(TObject *Sender)
{



	try{


	FormTestsRating->Visible=false;
	FormRecords->Visible=true;
	FormRecords->Show();
	}
	catch(...){}
}
//---------------------------------------------------------------------------
void __fastcall TFormTestsRating::FormClose(TObject *Sender, TCloseAction &Action)

{

	try{


	FormTestsRating->Visible=false;
	FormRecords->Visible=true;
	FormRecords->Show();
	}
	catch(...){}
}
//---------------------------------------------------------------------------


void __fastcall TFormTestsRating::ButtonAllClick(TObject *Sender)
{

//ExitMemo->Clear();
//ExitUsersMemo->Clear();
for (int i=0; i < StringGrid1->RowCount; i++) {
	StringGrid1->Rows[i]->Clear();
}
StringGrid1->Cells[0][0]="Результат";
StringGrid1->Cells[1][0]="Пользователь";
			AnsiString temp;
			AnsiString WayToFile;
			Variant ExcelApplication,ExcelBooks,Sheet,Cell;
			WayToFile=CURRENT_DIRECTORY+"\\Users.xlsx";
			int rowsCount;
			//int min, max;
			float mas[1000];
			AnsiString name[1000];
			AnsiString surname[1000];
			int i, choise ;

			if (ComboBoxChoise->ItemIndex == 0) {   choise  = 11;}
			else if (ComboBoxChoise->ItemIndex == 1) {    choise =21; }
			else if (ComboBoxChoise->ItemIndex == 2) {   choise =31;  }
			else if (ComboBoxChoise->ItemIndex == 3) {   choise =41;  }
			else  {  choise = 51 ;   }

			try{
				ExcelApplication=CreateOleObject("Excel.Application");
				ExcelBooks=ExcelApplication.OlePropertyGet("Workbooks").OlePropertyGet("Open",StringToOleStr(WayToFile));
				Sheet=ExcelBooks.OlePropertyGet("Worksheets",1);
				rowsCount=Sheet.OlePropertyGet("UsedRange").OlePropertyGet("Rows").OlePropertyGet("Count");
			}
			catch(...){
				Application->Title = "Ошибка";
				ShowMessage("Ошибка при обращении к файлу\n"+WayToFile);
				ExcelApplication.OleProcedure("Quit");
				return;
			}

			for (int i = 0; i < rowsCount; i++) {



			 temp=Sheet.OlePropertyGet("Cells").OlePropertyGet("Item",i+1,choise);
			 mas[i] = StrToFloat(temp);
			 temp=Sheet.OlePropertyGet("Cells").OlePropertyGet("Item",i+1,1);
			 name[i] = temp;
			 temp=Sheet.OlePropertyGet("Cells").OlePropertyGet("Item",i+1,5);
			 surname[i] = temp;

			}

		  //ExitMemo->Clear();
		  //ExitUsersMemo->Clear();

		StringGrid1->RowCount=rowsCount;
		float kol=0;
		AnsiString kolname = "nothing";
		AnsiString kolsurname = "nothing";
		for(int i = 0; i < rowsCount; i++)
		{
			for(int j = i + 1; j < rowsCount; j++)
			{
				if (mas[i] <mas[j])
				{
					kol=mas[i];
					kolname = name[i] ;
					kolsurname = surname[i] ;

					mas[i]=mas[j];
					name[i] = name[j];
					surname[i] = surname[j];

					mas[j]=kol;
					name[j] = kolname;
					surname[j] = kolsurname;
				}

			}
			//ExitMemo->Lines->Add(float(int(mas[i]*1000+0.5))/1000);
			//ExitUsersMemo->Lines->Add(name[i] + " " + surname[i]);
			StringGrid1->Cells[0][i+1]=FloatToStr(float(int(mas[i]*1000+0.5))/1000);
			StringGrid1->Cells[1][i+1]=name[i] + " " + surname[i];
		}


				//ExitMemo->Perform(EM_SCROLL,SB_LINEUP,0);
				//ExitUsersMemo->Perform(EM_SCROLL,SB_LINEUP,0);
				//ExcelApplication.OlePropertyGet("Workbooks").OlePropertyGet("Item",1).OleProcedure("Save");
				ExcelApplication.OleProcedure("Quit");

}
//---------------------------------------------------------------------------

void __fastcall TFormTestsRating::ButtonOldClick(TObject *Sender)
{

//ExitMemo->Clear();
//ExitUsersMemo->Clear();
for (int i=0; i < StringGrid1->RowCount; i++) {
	StringGrid1->Rows[i]->Clear();
}
StringGrid1->Cells[0][0]="Результат";
StringGrid1->Cells[1][0]="Пользователь";
AnsiString temp;
			AnsiString WayToFile;
			Variant ExcelApplication,ExcelBooks,Sheet,Cell;
			WayToFile=CURRENT_DIRECTORY+"\\Users.xlsx";
			int rowsCount;
			//int min, max;
			float mas[1000];
			AnsiString name[1000];
			AnsiString surname[1000];
			int i, choise ;

			if (ComboBoxChoise->ItemIndex == 0) {   choise  = 12;}
			else if (ComboBoxChoise->ItemIndex == 1) {    choise =22; }
				else if (ComboBoxChoise->ItemIndex == 2) {   choise =32;  }
			else if (ComboBoxChoise->ItemIndex == 3) {   choise =42;  }
			else  {  choise = 52 ;   }


			try{
				ExcelApplication=CreateOleObject("Excel.Application");
				ExcelBooks=ExcelApplication.OlePropertyGet("Workbooks").OlePropertyGet("Open",StringToOleStr(WayToFile));
				Sheet=ExcelBooks.OlePropertyGet("Worksheets",1);
				rowsCount=Sheet.OlePropertyGet("UsedRange").OlePropertyGet("Rows").OlePropertyGet("Count");
			}
			catch(...){
				Application->Title = "Ошибка";
				ShowMessage("Ошибка при обращении к файлу\n"+WayToFile);
				ExcelApplication.OleProcedure("Quit");
				return;
			}

			for (int i = 0; i < rowsCount; i++) {



			 temp=Sheet.OlePropertyGet("Cells").OlePropertyGet("Item",i+1,choise);
			 mas[i] = StrToFloat(temp);
			 temp=Sheet.OlePropertyGet("Cells").OlePropertyGet("Item",i+1,1);
			 name[i] = temp;
			 temp=Sheet.OlePropertyGet("Cells").OlePropertyGet("Item",i+1,5);
			 surname[i] = temp;

			}

		  //ExitMemo->Clear();
		  //ExitUsersMemo->Clear();

  {
		float kol=0;
		AnsiString kolname = "nothing";
		AnsiString kolsurname = "nothing";
		  for(int i = 0; i < rowsCount; i++)
	{
	   for(int j = i + 1; j < rowsCount; j++)
	  {
		if (mas[i] <mas[j])
		{
			kol=mas[i];
			kolname = name[i] ;
			kolsurname = surname[i] ;

			mas[i]=mas[j];
			name[i] = name[j];
			surname[i] = surname[j];

			mas[j]=kol;
			name[j] = kolname;
			surname[j] = kolsurname;
		}

	  }
	  //ExitMemo->Lines->Add(float(int(mas[i]*1000+0.5))/1000);
	  //ExitUsersMemo->Lines->Add(name[i] + " " + surname[i]);
	  StringGrid1->Cells[0][i+1]=FloatToStr(float(int(mas[i]*1000+0.5))/1000);
			StringGrid1->Cells[1][i+1]=name[i] + " " + surname[i];
	}
  }

				 //ExitMemo->Perform(EM_SCROLL,SB_LINEUP,0);
				 //ExitUsersMemo->Perform(EM_SCROLL,SB_LINEUP,0);
				//ExcelApplication.OlePropertyGet("Workbooks").OlePropertyGet("Item",1).OleProcedure("Save");
				ExcelApplication.OleProcedure("Quit");
}
//---------------------------------------------------------------------------

void __fastcall TFormTestsRating::ButtonRPClick(TObject *Sender)
{

//ExitMemo->Clear();
//ExitUsersMemo->Clear();
for (int i=0; i < StringGrid1->RowCount; i++) {
	StringGrid1->Rows[i]->Clear();
}
StringGrid1->Cells[0][0]="Результат";
StringGrid1->Cells[1][0]="Пользователь";
AnsiString temp;
			AnsiString WayToFile;
			Variant ExcelApplication,ExcelBooks,Sheet,Cell;
			WayToFile=CURRENT_DIRECTORY+"\\Users.xlsx";
			int rowsCount;
			//int min, max;
			float mas[1000];
			AnsiString name[1000];
			AnsiString surname[1000];
			int i, choise ;

			if (ComboBoxChoise->ItemIndex == 0) {   choise  = 14;}
			else if (ComboBoxChoise->ItemIndex == 1) {    choise =24; }
				else if (ComboBoxChoise->ItemIndex == 2) {   choise =34;  }
			else if (ComboBoxChoise->ItemIndex == 3) {   choise =44;  }
			else  {  choise = 54 ;   }

			try{
				ExcelApplication=CreateOleObject("Excel.Application");
				ExcelBooks=ExcelApplication.OlePropertyGet("Workbooks").OlePropertyGet("Open",StringToOleStr(WayToFile));
				Sheet=ExcelBooks.OlePropertyGet("Worksheets",1);
				rowsCount=Sheet.OlePropertyGet("UsedRange").OlePropertyGet("Rows").OlePropertyGet("Count");
			}
			catch(...){
				Application->Title = "Ошибка";
				ShowMessage("Ошибка при обращении к файлу\n"+WayToFile);
				ExcelApplication.OleProcedure("Quit");
				return;
			}

			for (int i = 0; i < rowsCount; i++) {



			 temp=Sheet.OlePropertyGet("Cells").OlePropertyGet("Item",i+1,choise);
			 mas[i] = StrToFloat(temp);
			 temp=Sheet.OlePropertyGet("Cells").OlePropertyGet("Item",i+1,1);
			 name[i] = temp;
			 temp=Sheet.OlePropertyGet("Cells").OlePropertyGet("Item",i+1,5);
			 surname[i] = temp;

			}

		  //ExitMemo->Clear();
		  //ExitUsersMemo->Clear();

  {
		float kol=0;
		AnsiString kolname = "nothing";
		AnsiString kolsurname = "nothing";
		  for(int i = 0; i < rowsCount; i++)
	{
	   for(int j = i + 1; j < rowsCount; j++)
	  {
		if (mas[i] <mas[j])
		{
			kol=mas[i];
			kolname = name[i] ;
			kolsurname = surname[i] ;

			mas[i]=mas[j];
			name[i] = name[j];
			surname[i] = surname[j];

			mas[j]=kol;
			name[j] = kolname;
			surname[j] = kolsurname;
		}

	  }
	  //ExitMemo->Lines->Add(float(int(mas[i]*1000+0.5))/1000);
	  //ExitUsersMemo->Lines->Add(name[i] + " " + surname[i]);
	  StringGrid1->Cells[0][i+1]=FloatToStr(float(int(mas[i]*1000+0.5))/1000);
			StringGrid1->Cells[1][i+1]=name[i] + " " + surname[i];
	}
  }

				 //ExitMemo->Perform(EM_SCROLL,SB_LINEUP,0);
				 //ExitUsersMemo->Perform(EM_SCROLL,SB_LINEUP,0);
				//ExcelApplication.OlePropertyGet("Workbooks").OlePropertyGet("Item",1).OleProcedure("Save");
				ExcelApplication.OleProcedure("Quit");
}
//---------------------------------------------------------------------------

void __fastcall TFormTestsRating::ButtonRIClick(TObject *Sender)
{
//ComboBoxChoise->Enabled = true;
//ExitMemo->Clear();
//ExitUsersMemo->Clear();
for (int i=0; i < StringGrid1->RowCount; i++) {
	StringGrid1->Rows[i]->Clear();
}
StringGrid1->Cells[0][0]="Результат";
StringGrid1->Cells[1][0]="Пользователь";
AnsiString temp;
			AnsiString WayToFile;
			Variant ExcelApplication,ExcelBooks,Sheet,Cell;
			WayToFile=CURRENT_DIRECTORY+"\\Users.xlsx";
			int rowsCount;
			//int min, max;
			float mas[1000];
			AnsiString name[1000];
			AnsiString surname[1000];
			int i, choise ;

			if (ComboBoxChoise->ItemIndex == 0) {   choise  = 15;}
			else if (ComboBoxChoise->ItemIndex == 1) {    choise =25; }
			else if (ComboBoxChoise->ItemIndex == 2) {   choise =35;  }
			else if (ComboBoxChoise->ItemIndex == 3) {   choise =45;  }
			else  {  choise = 55 ;   }

			try{
				ExcelApplication=CreateOleObject("Excel.Application");
				ExcelBooks=ExcelApplication.OlePropertyGet("Workbooks").OlePropertyGet("Open",StringToOleStr(WayToFile));
				Sheet=ExcelBooks.OlePropertyGet("Worksheets",1);
				rowsCount=Sheet.OlePropertyGet("UsedRange").OlePropertyGet("Rows").OlePropertyGet("Count");
			}
			catch(...){
				Application->Title = "Ошибка";
				ShowMessage("Ошибка при обращении к файлу\n"+WayToFile);
				ExcelApplication.OleProcedure("Quit");
				return;
			}

			for (int i = 0; i < rowsCount; i++) {



			 temp=Sheet.OlePropertyGet("Cells").OlePropertyGet("Item",i+1,choise);
			 mas[i] = StrToFloat(temp);
			 temp=Sheet.OlePropertyGet("Cells").OlePropertyGet("Item",i+1,1);
			 name[i] = temp;
			 temp=Sheet.OlePropertyGet("Cells").OlePropertyGet("Item",i+1,5);
			 surname[i] = temp;

			}

		  //ExitMemo->Clear();
		  //ExitUsersMemo->Clear();

  {
		float kol=0;
		AnsiString kolname = "nothing";
		AnsiString kolsurname = "nothing";
		  for(int i = 0; i < rowsCount; i++)
	{
	   for(int j = i + 1; j < rowsCount; j++)
	  {
		if (mas[i] <mas[j])
		{
			kol=mas[i];
			kolname = name[i] ;
			kolsurname = surname[i] ;

			mas[i]=mas[j];
			name[i] = name[j];
			surname[i] = surname[j];

			mas[j]=kol;
			name[j] = kolname;
			surname[j] = kolsurname;
		}

	  }
	  //ExitMemo->Lines->Add(float(int(mas[i]*1000+0.5))/1000);
	  //ExitUsersMemo->Lines->Add(name[i] + " " + surname[i]);
	  StringGrid1->Cells[0][i+1]=FloatToStr(float(int(mas[i]*1000+0.5))/1000);
			StringGrid1->Cells[1][i+1]=name[i] + " " + surname[i];
	}
  }

				//ExitMemo->Perform(EM_SCROLL,SB_LINEUP,0);
				// ExitUsersMemo->Perform(EM_SCROLL,SB_LINEUP,0);
				//ExcelApplication.OlePropertyGet("Workbooks").OlePropertyGet("Item",1).OleProcedure("Save");
				ExcelApplication.OleProcedure("Quit");
}
//---------------------------------------------------------------------------

void __fastcall TFormTestsRating::ButtonVKLClick(TObject *Sender)
{
//ComboBoxChoise->Enabled = true;
//ExitMemo->Clear();
//ExitUsersMemo->Clear();
for (int i=0; i < StringGrid1->RowCount; i++) {
	StringGrid1->Rows[i]->Clear();
}
StringGrid1->Cells[0][0]="Результат";
StringGrid1->Cells[1][0]="Пользователь";
AnsiString temp;
			AnsiString WayToFile;
			Variant ExcelApplication,ExcelBooks,Sheet,Cell;
			WayToFile=CURRENT_DIRECTORY+"\\Users.xlsx";
			int rowsCount;
			//int min, max;
			float mas[1000];
			AnsiString name[1000];
			AnsiString surname[1000];
			int i, choise ;

			if (ComboBoxChoise->ItemIndex == 0) {   choise  = 13;}
			else if (ComboBoxChoise->ItemIndex == 1) {    choise =23; }
			else if (ComboBoxChoise->ItemIndex == 2) {   choise =33;  }
			else if (ComboBoxChoise->ItemIndex == 3) {   choise =43;  }
			else  {  choise = 53 ;   }

			try{
				ExcelApplication=CreateOleObject("Excel.Application");
				ExcelBooks=ExcelApplication.OlePropertyGet("Workbooks").OlePropertyGet("Open",StringToOleStr(WayToFile));
				Sheet=ExcelBooks.OlePropertyGet("Worksheets",1);
				rowsCount=Sheet.OlePropertyGet("UsedRange").OlePropertyGet("Rows").OlePropertyGet("Count");
			}
			catch(...){
				Application->Title = "Ошибка";
				ShowMessage("Ошибка при обращении к файлу\n"+WayToFile);
				ExcelApplication.OleProcedure("Quit");
				return;
			}

			for (int i = 0; i < rowsCount; i++) {



			 temp=Sheet.OlePropertyGet("Cells").OlePropertyGet("Item",i+1,choise);
			 mas[i] = StrToFloat(temp);
			 temp=Sheet.OlePropertyGet("Cells").OlePropertyGet("Item",i+1,1);
			 name[i] = temp;
			 temp=Sheet.OlePropertyGet("Cells").OlePropertyGet("Item",i+1,5);
			 surname[i] = temp;

			}

		  //ExitMemo->Clear();
		  //ExitUsersMemo->Clear();

  {
		float kol=0;
		AnsiString kolname = "nothing";
		AnsiString kolsurname = "nothing";
		  for(int i = 0; i < rowsCount; i++)
	{
	   for(int j = i + 1; j < rowsCount; j++)
	  {
		if (mas[i] <mas[j])
		{
			kol=mas[i];
			kolname = name[i] ;
			kolsurname = surname[i] ;

			mas[i]=mas[j];
			name[i] = name[j];
			surname[i] = surname[j];

			mas[j]=kol;
			name[j] = kolname;
			surname[j] = kolsurname;
		}

	  }
	  //ExitMemo->Lines->Add(float(int(mas[i]*1000+0.5))/1000);
	  //ExitUsersMemo->Lines->Add(name[i] + " " + surname[i]);
	  StringGrid1->Cells[0][i+1]=FloatToStr(float(int(mas[i]*1000+0.5))/1000);
			StringGrid1->Cells[1][i+1]=name[i] + " " + surname[i];
	}
  }

				//ExitMemo->Perform(EM_SCROLL,SB_LINEUP,0);
				 //ExitUsersMemo->Perform(EM_SCROLL,SB_LINEUP,0);
				//ExcelApplication.OlePropertyGet("Workbooks").OlePropertyGet("Item",1).OleProcedure("Save");
				ExcelApplication.OleProcedure("Quit");
}
//---------------------------------------------------------------------------

void __fastcall TFormTestsRating::ButtonNoWarClick(TObject *Sender)
{
//ComboBoxChoise->Enabled = true;
//ExitMemo->Clear();
//ExitUsersMemo->Clear();
for (int i=0; i < StringGrid1->RowCount; i++) {
	StringGrid1->Rows[i]->Clear();
}
StringGrid1->Cells[0][0]="Результат";
StringGrid1->Cells[1][0]="Пользователь";
AnsiString temp;
			AnsiString WayToFile;
			Variant ExcelApplication,ExcelBooks,Sheet,Cell;
			WayToFile=CURRENT_DIRECTORY+"\\Users.xlsx";
			int rowsCount;
			//int min, max;
			float mas[1000];
			AnsiString name[1000];
			AnsiString surname[1000];
			int i, choise ;

			if (ComboBoxChoise->ItemIndex == 0) {   choise  = 17;}
			else if (ComboBoxChoise->ItemIndex == 1) {    choise =27; }
			else if (ComboBoxChoise->ItemIndex == 2) {   choise =37;  }
			else if (ComboBoxChoise->ItemIndex == 3) {   choise =47;  }
			else  {  choise = 57 ;   }

			try{
				ExcelApplication=CreateOleObject("Excel.Application");
				ExcelBooks=ExcelApplication.OlePropertyGet("Workbooks").OlePropertyGet("Open",StringToOleStr(WayToFile));
				Sheet=ExcelBooks.OlePropertyGet("Worksheets",1);
				rowsCount=Sheet.OlePropertyGet("UsedRange").OlePropertyGet("Rows").OlePropertyGet("Count");
			}
			catch(...){
				Application->Title = "Ошибка";
				ShowMessage("Ошибка при обращении к файлу\n"+WayToFile);
				ExcelApplication.OleProcedure("Quit");
				return;
			}

			for (int i = 0; i < rowsCount; i++) {



			 temp=Sheet.OlePropertyGet("Cells").OlePropertyGet("Item",i+1,choise);
			 mas[i] = StrToFloat(temp);
			 temp=Sheet.OlePropertyGet("Cells").OlePropertyGet("Item",i+1,1);
			 name[i] = temp;
			 temp=Sheet.OlePropertyGet("Cells").OlePropertyGet("Item",i+1,5);
			 surname[i] = temp;

			}

		  //ExitMemo->Clear();
		  //ExitUsersMemo->Clear();

  {
		float kol=0;
		AnsiString kolname = "nothing";
		AnsiString kolsurname = "nothing";
		  for(int i = 0; i < rowsCount; i++)
	{
	   for(int j = i + 1; j < rowsCount; j++)
	  {
		if (mas[i] <mas[j])
		{
			kol=mas[i];
			kolname = name[i] ;
			kolsurname = surname[i] ;

			mas[i]=mas[j];
			name[i] = name[j];
			surname[i] = surname[j];

			mas[j]=kol;
			name[j] = kolname;
			surname[j] = kolsurname;
		}

	  }
	  //ExitMemo->Lines->Add(float(int(mas[i]*1000+0.5))/1000);
	  //ExitUsersMemo->Lines->Add(name[i] + " " + surname[i]);
	  StringGrid1->Cells[0][i+1]=FloatToStr(float(int(mas[i]*1000+0.5))/1000);
			StringGrid1->Cells[1][i+1]=name[i] + " " + surname[i];
	}
  }

				 //ExitMemo->Perform(EM_SCROLL,SB_LINEUP,0);
				 //ExitUsersMemo->Perform(EM_SCROLL,SB_LINEUP,0);
				//ExcelApplication.OlePropertyGet("Workbooks").OlePropertyGet("Item",1).OleProcedure("Save");
				ExcelApplication.OleProcedure("Quit");
}
//---------------------------------------------------------------------------

void __fastcall TFormTestsRating::ButtonAfterWarClick(TObject *Sender)
{
//ComboBoxChoise->Enabled = true;
//ExitMemo->Clear();
//ExitUsersMemo->Clear();
for (int i=0; i < StringGrid1->RowCount; i++) {
	StringGrid1->Rows[i]->Clear();
}
StringGrid1->Cells[0][0]="Результат";
StringGrid1->Cells[1][0]="Пользователь";
AnsiString temp;
			AnsiString WayToFile;
			Variant ExcelApplication,ExcelBooks,Sheet,Cell;
			WayToFile=CURRENT_DIRECTORY+"\\Users.xlsx";
			int rowsCount;
			//int min, max;
			float mas[1000];
			AnsiString name[1000];
			AnsiString surname[1000];
			int i, choise ;

			if (ComboBoxChoise->ItemIndex == 0) {   choise  = 19;}
			else if (ComboBoxChoise->ItemIndex == 1) {    choise =29; }
				else if (ComboBoxChoise->ItemIndex == 2) {   choise =39;  }
			else if (ComboBoxChoise->ItemIndex == 3) {   choise =49;  }
			else  {  choise = 59 ;   }

			try{
				ExcelApplication=CreateOleObject("Excel.Application");
				ExcelBooks=ExcelApplication.OlePropertyGet("Workbooks").OlePropertyGet("Open",StringToOleStr(WayToFile));
				Sheet=ExcelBooks.OlePropertyGet("Worksheets",1);
				rowsCount=Sheet.OlePropertyGet("UsedRange").OlePropertyGet("Rows").OlePropertyGet("Count");
			}
			catch(...){
				Application->Title = "Ошибка";
				ShowMessage("Ошибка при обращении к файлу\n"+WayToFile);
				ExcelApplication.OleProcedure("Quit");
				return;
			}

			for (int i = 0; i < rowsCount; i++) {



			 temp=Sheet.OlePropertyGet("Cells").OlePropertyGet("Item",i+1,choise);
			 mas[i] = StrToFloat(temp);
			 temp=Sheet.OlePropertyGet("Cells").OlePropertyGet("Item",i+1,1);
			 name[i] = temp;
			 temp=Sheet.OlePropertyGet("Cells").OlePropertyGet("Item",i+1,5);
			 surname[i] = temp;

			}

		  //ExitMemo->Clear();
		  //ExitUsersMemo->Clear();

  {
		float kol=0;
		AnsiString kolname = "nothing";
		AnsiString kolsurname = "nothing";
		  for(int i = 0; i < rowsCount; i++)
	{
	   for(int j = i + 1; j < rowsCount; j++)
	  {
		if (mas[i] <mas[j])
		{
			kol=mas[i];
			kolname = name[i] ;
			kolsurname = surname[i] ;

			mas[i]=mas[j];
			name[i] = name[j];
			surname[i] = surname[j];

			mas[j]=kol;
			name[j] = kolname;
			surname[j] = kolsurname;
		}

	  }
	  //ExitMemo->Lines->Add(float(int(mas[i]*1000+0.5))/1000);
	  //ExitUsersMemo->Lines->Add(name[i] + " " + surname[i]);
	  StringGrid1->Cells[0][i+1]=FloatToStr(float(int(mas[i]*1000+0.5))/1000);
			StringGrid1->Cells[1][i+1]=name[i] + " " + surname[i];
	}
  }

				 //ExitMemo->Perform(EM_SCROLL,SB_LINEUP,0);
				 //ExitUsersMemo->Perform(EM_SCROLL,SB_LINEUP,0);
				//ExcelApplication.OlePropertyGet("Workbooks").OlePropertyGet("Item",1).OleProcedure("Save");
				ExcelApplication.OleProcedure("Quit");
}
//---------------------------------------------------------------------------

void __fastcall TFormTestsRating::ButtonMediaClick(TObject *Sender)
{
//ComboBoxChoise->Enabled = true;
//ExitMemo->Clear();
//ExitUsersMemo->Clear();
for (int i=0; i < StringGrid1->RowCount; i++) {
	StringGrid1->Rows[i]->Clear();
}
StringGrid1->Cells[0][0]="Результат";
StringGrid1->Cells[1][0]="Пользователь";
AnsiString temp;
			AnsiString WayToFile;
			Variant ExcelApplication,ExcelBooks,Sheet,Cell;
			WayToFile=CURRENT_DIRECTORY+"\\Users.xlsx";
			int rowsCount;
			//int min, max;
			float mas[1000];
			AnsiString name[1000];
			AnsiString surname[1000];
			int i, choise ;

			if (ComboBoxChoise->ItemIndex == 0) {   choise  = 20;}
			else if (ComboBoxChoise->ItemIndex == 1) {    choise =30; }
				else if (ComboBoxChoise->ItemIndex == 2) {   choise =40;  }
			else if (ComboBoxChoise->ItemIndex == 3) {   choise =50;  }
			else  {  choise = 60 ;   }

			try{
				ExcelApplication=CreateOleObject("Excel.Application");
				ExcelBooks=ExcelApplication.OlePropertyGet("Workbooks").OlePropertyGet("Open",StringToOleStr(WayToFile));
				Sheet=ExcelBooks.OlePropertyGet("Worksheets",1);
				rowsCount=Sheet.OlePropertyGet("UsedRange").OlePropertyGet("Rows").OlePropertyGet("Count");
			}
			catch(...){
				Application->Title = "Ошибка";
				ShowMessage("Ошибка при обращении к файлу\n"+WayToFile);
				ExcelApplication.OleProcedure("Quit");
				return;
			}

			for (int i = 0; i < rowsCount; i++) {



			 temp=Sheet.OlePropertyGet("Cells").OlePropertyGet("Item",i+1,choise);
			 mas[i] = StrToFloat(temp);
			 temp=Sheet.OlePropertyGet("Cells").OlePropertyGet("Item",i+1,1);
			 name[i] = temp;
			 temp=Sheet.OlePropertyGet("Cells").OlePropertyGet("Item",i+1,5);
			 surname[i] = temp;

			}

		  //ExitMemo->Clear();
		  //ExitUsersMemo->Clear();

  {
		float kol=0;
		AnsiString kolname = "nothing";
		AnsiString kolsurname = "nothing";
		  for(int i = 0; i < rowsCount; i++)
	{
	   for(int j = i + 1; j < rowsCount; j++)
	  {
		if (mas[i] <mas[j])
		{
			kol=mas[i];
			kolname = name[i] ;
			kolsurname = surname[i] ;

			mas[i]=mas[j];
			name[i] = name[j];
			surname[i] = surname[j];

			mas[j]=kol;
			name[j] = kolname;
			surname[j] = kolsurname;
		}

	  }
	  //ExitMemo->Lines->Add(float(int(mas[i]*1000+0.5))/1000);
	  //ExitUsersMemo->Lines->Add(name[i] + " " + surname[i]);
	  StringGrid1->Cells[0][i+1]=FloatToStr(float(int(mas[i]*1000+0.5))/1000);
			StringGrid1->Cells[1][i+1]=name[i] + " " + surname[i];
	}
  }

				 //ExitMemo->Perform(EM_SCROLL,SB_LINEUP,0);
				 //ExitUsersMemo->Perform(EM_SCROLL,SB_LINEUP,0);
				//ExcelApplication.OlePropertyGet("Workbooks").OlePropertyGet("Item",1).OleProcedure("Save");
				ExcelApplication.OleProcedure("Quit");
}
//---------------------------------------------------------------------------

void __fastcall TFormTestsRating::ButtonTwentyClick(TObject *Sender)
{
//ComboBoxChoise->Enabled = true;
//ExitMemo->Clear();
//ExitUsersMemo->Clear();
for (int i=0; i < StringGrid1->RowCount; i++) {
	StringGrid1->Rows[i]->Clear();
}
StringGrid1->Cells[0][0]="Результат";
StringGrid1->Cells[1][0]="Пользователь";
AnsiString temp;
			AnsiString WayToFile;
			Variant ExcelApplication,ExcelBooks,Sheet,Cell;
			WayToFile=CURRENT_DIRECTORY+"\\Users.xlsx";
			int rowsCount;
			//int min, max;
			float mas[1000];
			AnsiString name[1000];
			AnsiString surname[1000];
			int i, choise ;

			if (ComboBoxChoise->ItemIndex == 0) {   choise  = 16;}
			else if (ComboBoxChoise->ItemIndex == 1) {    choise =26; }
			else if (ComboBoxChoise->ItemIndex == 2) {   choise =36;  }
			else if (ComboBoxChoise->ItemIndex == 3) {   choise =46;  }
			else  {  choise = 56 ;   }

			try{
				ExcelApplication=CreateOleObject("Excel.Application");
				ExcelBooks=ExcelApplication.OlePropertyGet("Workbooks").OlePropertyGet("Open",StringToOleStr(WayToFile));
				Sheet=ExcelBooks.OlePropertyGet("Worksheets",1);
				rowsCount=Sheet.OlePropertyGet("UsedRange").OlePropertyGet("Rows").OlePropertyGet("Count");
			}
			catch(...){
				Application->Title = "Ошибка";
				ShowMessage("Ошибка при обращении к файлу\n"+WayToFile);
				ExcelApplication.OleProcedure("Quit");
				return;
			}

			for (int i = 0; i < rowsCount; i++) {



			 temp=Sheet.OlePropertyGet("Cells").OlePropertyGet("Item",i+1,choise);
			 mas[i] = StrToFloat(temp);
			 temp=Sheet.OlePropertyGet("Cells").OlePropertyGet("Item",i+1,1);
			 name[i] = temp;
			 temp=Sheet.OlePropertyGet("Cells").OlePropertyGet("Item",i+1,5);
			 surname[i] = temp;

			}

		  //ExitMemo->Clear();
		  //ExitUsersMemo->Clear();

  {
		float kol=0;
		AnsiString kolname = "nothing";
		AnsiString kolsurname = "nothing";
		  for(int i = 0; i < rowsCount; i++)
	{
	   for(int j = i + 1; j < rowsCount; j++)
	  {
		if (mas[i] <mas[j])
		{
			kol=mas[i];
			kolname = name[i] ;
			kolsurname = surname[i] ;

			mas[i]=mas[j];
			name[i] = name[j];
			surname[i] = surname[j];

			mas[j]=kol;
			name[j] = kolname;
			surname[j] = kolsurname;
		}

	  }
	  //ExitMemo->Lines->Add(float(int(mas[i]*1000+0.5))/1000);
	  //ExitUsersMemo->Lines->Add(name[i] + " " + surname[i]);
	  StringGrid1->Cells[0][i+1]=FloatToStr(float(int(mas[i]*1000+0.5))/1000);
			StringGrid1->Cells[1][i+1]=name[i] + " " + surname[i];
	}
  }

				// ExitMemo->Perform(EM_SCROLL,SB_LINEUP,0);
				// ExitUsersMemo->Perform(EM_SCROLL,SB_LINEUP,0);
				//ExcelApplication.OlePropertyGet("Workbooks").OlePropertyGet("Item",1).OleProcedure("Save");
				ExcelApplication.OleProcedure("Quit");
}
//---------------------------------------------------------------------------

void __fastcall TFormTestsRating::ButtonWARClick(TObject *Sender)
{
//ComboBoxChoise->Enabled = true;
//ExitMemo->Clear();
//ExitUsersMemo->Clear();
for (int i=0; i < StringGrid1->RowCount; i++) {
	StringGrid1->Rows[i]->Clear();
}
StringGrid1->Cells[0][0]="Результат";
StringGrid1->Cells[1][0]="Пользователь";
AnsiString temp;
			AnsiString WayToFile;
			Variant ExcelApplication,ExcelBooks,Sheet,Cell;
			WayToFile=CURRENT_DIRECTORY+"\\Users.xlsx";
			int rowsCount;
			//int min, max;
			float mas[1000];
			AnsiString name[1000];
			AnsiString surname[1000];
			int i, choise ;

			if (ComboBoxChoise->ItemIndex == 0) {   choise  = 18;}
			else if (ComboBoxChoise->ItemIndex == 1) {    choise =28; }
			else if (ComboBoxChoise->ItemIndex == 2) {   choise =38;  }
			else if (ComboBoxChoise->ItemIndex == 3) {   choise =48;  }
			else  {  choise = 58 ;   }

			try{
				ExcelApplication=CreateOleObject("Excel.Application");
				ExcelBooks=ExcelApplication.OlePropertyGet("Workbooks").OlePropertyGet("Open",StringToOleStr(WayToFile));
				Sheet=ExcelBooks.OlePropertyGet("Worksheets",1);
				rowsCount=Sheet.OlePropertyGet("UsedRange").OlePropertyGet("Rows").OlePropertyGet("Count");
			}
			catch(...){
				Application->Title = "Ошибка";
				ShowMessage("Ошибка при обращении к файлу\n"+WayToFile);
				ExcelApplication.OleProcedure("Quit");
				return;
			}

			for (int i = 0; i < rowsCount; i++) {



			 temp=Sheet.OlePropertyGet("Cells").OlePropertyGet("Item",i+1,choise);
			 mas[i] = StrToFloat(temp);
			 temp=Sheet.OlePropertyGet("Cells").OlePropertyGet("Item",i+1,1);
			 name[i] = temp;
			 temp=Sheet.OlePropertyGet("Cells").OlePropertyGet("Item",i+1,5);
			 surname[i] = temp;

			}

		 // ExitMemo->Clear();
		  //ExitUsersMemo->Clear();

  {
		float kol=0;
		AnsiString kolname = "nothing";
		AnsiString kolsurname = "nothing";
		  for(int i = 0; i < rowsCount; i++)
	{
	   for(int j = i + 1; j < rowsCount; j++)
	  {
		if (mas[i] <mas[j])
		{
			kol=mas[i];
			kolname = name[i] ;
			kolsurname = surname[i] ;

			mas[i]=mas[j];
			name[i] = name[j];
			surname[i] = surname[j];

			mas[j]=kol;
			name[j] = kolname;
			surname[j] = kolsurname;
		}

	  }
	  //ExitMemo->Lines->Add(float(int(mas[i]*1000+0.5))/1000);
	  //ExitUsersMemo->Lines->Add(name[i] + " " + surname[i]);
	  StringGrid1->Cells[0][i+1]=FloatToStr(float(int(mas[i]*1000+0.5))/1000);
	  StringGrid1->Cells[1][i+1]=name[i] + " " + surname[i];
	}
  }

				// ExitMemo->Perform(EM_SCROLL,SB_LINEUP,0);
				// ExitUsersMemo->Perform(EM_SCROLL,SB_LINEUP,0);
				//ExcelApplication.OlePropertyGet("Workbooks").OlePropertyGet("Item",1).OleProcedure("Save");
				ExcelApplication.OleProcedure("Quit");
}
//---------------------------------------------------------------------------


void __fastcall TFormTestsRating::ComboBoxChoiseChange(TObject *Sender)
{
for (int i=0; i < StringGrid1->RowCount; i++) {
	StringGrid1->Rows[i]->Clear();
}
StringGrid1->Cells[0][0]="Результат";
StringGrid1->Cells[1][0]="Пользователь";

ButtonAll->Enabled = true;
ButtonOld->Enabled = true;
ButtonRP->Enabled = true;
ButtonRI->Enabled = true;
ButtonVKL->Enabled = true;
ButtonNoWar->Enabled = true;
ButtonAfterWar->Enabled = true;
ButtonMedia->Enabled = true;
ButtonTwenty->Enabled = true;
ButtonWAR->Enabled = true;

}
//---------------------------------------------------------------------------

void __fastcall TFormTestsRating::FormActivate(TObject *Sender)
{
//ExitMemo->Clear();
//ExitUsersMemo->Clear();
}
//---------------------------------------------------------------------------

void __fastcall TFormTestsRating::ScrollBar1Scroll(TObject *Sender, TScrollCode ScrollCode,
          int &ScrollPos)
{
	//ExitMemo->Perform(EM_SCROLL,SB_LINEDOWN,0);
	//ExitUsersMemo->Perform(EM_SCROLL,SB_LINEDOWN,0);
}
//---------------------------------------------------------------------------


