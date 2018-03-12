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
#include "UnitTest.h"
#include "UnitUser.h"

#include <stdio.h>
#include <io.h>
#include <direct.h>
#include <ComObj.hpp>
#include <utilcls.h>

//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma resource "*.dfm"
TFStart *FStart;

//---------------------------------------------------------------------------
__fastcall TFStart::TFStart(TComponent* Owner)
	: TForm(Owner)
{
}

bool wasStretched=false;
bool wasNodeAdded;
bool wasParametresChanged=false;
bool wasSettingsChanged=false;
bool areImagesInFolder;
bool areImagesInCreated;
bool areMediaInFolder;
bool wasChecked=true;

bool isRandomTest;
AnsiString username;
AnsiString usersurname;
AnsiString usergroup;

AnsiString CURRENT_DIRECTORY;

int resolutionX,resolutionY;
const int NORMAL_RESOLUTION_X=1920;
const int NORMAL_RESOLUTION_Y=1080;
int LINE_LENGTH=50;
int IMAGE_WIDTH=350;
int IMAGE_HEIGHT=250;
int MAX_IMAGE_HEIGHT=700;
int MAX_IMAGE_WIDTH=1200;
int WIDTH_INFO=300;
int WIDTH_ANSWERS=1000;
int HEIGHT=50;
int FONT_HEIGHT=32;
int INDENTION=10;
int ADDITION_WIDTH=500;
int FORM_HEIGHT=900;
int FORM_WIDTH=1700;

extern int SECONDS_FOR_ANSWER;

void screenParametres(){
		double dx=(double)resolutionX / NORMAL_RESOLUTION_X;
		double dy=(double)resolutionY / NORMAL_RESOLUTION_Y;
		LINE_LENGTH*=dx;
		IMAGE_WIDTH*=dx;
		IMAGE_HEIGHT*=dy;
		WIDTH_INFO*=dx;
		WIDTH_ANSWERS*=dx;
		HEIGHT*=dy;
		MAX_IMAGE_HEIGHT*=dy;
		MAX_IMAGE_WIDTH*=dx;
		FONT_HEIGHT*=dy;
		ADDITION_WIDTH*=dx;
		FORM_HEIGHT*=dy;
		FORM_WIDTH*=dx;
		if (dx>dy) INDENTION*=dx;
		else  INDENTION*=dy;
}

void getFiles(AnsiString SubDirectory, TTreeView* TreeView, AnsiString Extension, bool* wasAdded){
	struct _finddata_t fileNames;
	intptr_t file;     int i;
	AnsiString directory;

	try{
		directory=CURRENT_DIRECTORY+SubDirectory;
		chdir(directory.c_str());
	}
	catch(...){
		Application->Title="Ошибка";
		ShowMessage("Ошибка при входе в директорию "+directory);
	}

	try{
		TreeView->Items->Clear();
	}
	catch(...){
		Application->Title="Ошибка";
		ShowMessage("Ошибка компонента TreeView");
	}
	try{
		AnsiString Format="*"+Extension;
		if ((file=_findfirst(Format.c_str(),&fileNames))!=-1) {
			*wasAdded=true;
			do{
				UnicodeString name=(UnicodeString)fileNames.name;
				for (i = 0; i < Extension.Length(); i++) {
					name.Delete(name.Length(),1);
				}
				TreeView->Items->Add(NULL,name);
			} while( _findnext(file,&fileNames)==0) ;
		}
		else *wasAdded=false;
	}
	catch(...){
		Application->Title="Ошибка";
		ShowMessage("Ошибка при поиске файлов");
	}
	try{
		chdir(CURRENT_DIRECTORY.c_str());
	}
	catch(...){
		Application->Title="Ошибка";
		ShowMessage("Ошибка при входе в директорию "+CURRENT_DIRECTORY);
	}
}

int getBackgroundsAmount(){
	struct _finddata_t fileNames;
	intptr_t file;
	AnsiString directory;
	int amount=0;

	try{
		directory=CURRENT_DIRECTORY+"\\backgrounds";
		chdir(directory.c_str());
	}
	catch(...){
		return 0;
	}

	try{
		if ((file=_findfirst("*.jpg",&fileNames))!=-1) {
			amount++;
			do{
				amount++;
			} while( _findnext(file,&fileNames)==0) ;
		}
		else amount=0;
	}
	catch(...){
		return 0;
	}
	try{
		chdir(CURRENT_DIRECTORY.c_str());
	}
	catch(...){
		Application->Title="Ошибка";
		ShowMessage("Ошибка при входе в директорию "+CURRENT_DIRECTORY+"\nПрограмма будет завершена");
		FStart->Close();
	}
	return amount;
}

void open(){
	FStart->Hide();
	FTest->Show();
	FTest->Visible=true;
	FTest->OnActivate;
}

void __fastcall TFStart::ButtonStartRandomTestClick(TObject *Sender)
{
	username = ComboBoxUsers->Text;
	if (FCharacteristics->ComboBoxTheme->Text == "Мультимедийные вопросы" && FCharacteristics->CheckBoxMedia->Checked==false ) {
		FCharacteristics->CheckBoxMedia->Checked = true;
	}
	FCharacteristics->OnActivate(Sender);

	isRandomTest=true;
	open();
}
//---------------------------------------------------------------------------

void __fastcall TFStart::ButtonSettingsClick(TObject *Sender)
{
	FCharacteristics->Show();
	FStart->Hide();
}
//---------------------------------------------------------------------------

void __fastcall TFStart::ButtonStartCreatedTestClick(TObject *Sender)
{
	try{
		getFiles("\\Testings",FStart->TreeViewTestings,".xlsx",&wasNodeAdded);
	}
	catch(...){
		ShowMessage("s");
	}
	if (wasStretched==false) {

		BevelFrame->Left=FStart->ClientWidth;
		BevelFrame->Top=INDENTION/2;
		BevelFrame->Height=FStart->ClientHeight-INDENTION;
		BevelFrame->Width=ADDITION_WIDTH - INDENTION/2;

		LabelTestings->Top=BevelFrame->Top +INDENTION/2;
		LabelTestings->Left=BevelFrame->Left + (BevelFrame->Width - LabelTestings->Width)/2;

		TreeViewTestings->Left=BevelFrame->Left+INDENTION;
		TreeViewTestings->Top=LabelTestings->Top+LabelTestings->Height;
		TreeViewTestings->Width=BevelFrame->Width - 2*INDENTION;

        MemoStatistics->Left=TreeViewTestings->Left;
		MemoStatistics->Top=TreeViewTestings->Top + TreeViewTestings->Height + INDENTION;
		MemoStatistics->Width=TreeViewTestings->Width;
		MemoStatistics->Lines->Clear();

		ButtonStart->Left=TreeViewTestings->Left + (TreeViewTestings->Width - ButtonStart->Width - ButtonDeleteTest->Width - ButtonDelAll->Width -2*INDENTION)/2;
		ButtonStart->Top=MemoStatistics->Top+MemoStatistics->Height+ INDENTION;

		ButtonDeleteTest->Top=ButtonStart->Top;
		ButtonDeleteTest->Left=ButtonStart->Left+ButtonStart->Width+INDENTION;

		ButtonDelAll->Top=ButtonStart->Top;
		ButtonDelAll->Left=ButtonDeleteTest->Left+ButtonDeleteTest->Width+INDENTION;
		try{
			TreeViewTestings->Selected = TreeViewTestings->Items->GetFirstNode();
		}
		catch(...){
			MemoStatistics->Lines->Add("Нет сохраненных тестирований");
		}
		ButtonStartCreatedTest->Enabled=false;
		if (FStart->ClientWidth <=1000){
			TimerForOpening->Enabled=true;
			wasStretched=true;
		}
	}

	if (TreeViewTestings->Items->Count==0) {
		LabelTestings->Caption="Нет сохраненных тестирований";
		LabelTestings->Cursor=crHandPoint;
		LabelTestings->Top=(FStart->ClientHeight-LabelTestings->Height)/2;
		LabelTestings->Left=BevelFrame->Left + (BevelFrame->Width - LabelTestings->Width)/2;
		TreeViewTestings->Visible=false;
	}
	else {
		TreeViewTestings->Visible=true;
		LabelTestings->Cursor=crDefault;
		LabelTestings->Caption="Выберите тестирование";
		LabelTestings->Top=BevelFrame->Top +INDENTION/2;
		LabelTestings->Left=BevelFrame->Left + (BevelFrame->Width - LabelTestings->Width)/2;
	}
	BevelFrame->Visible=true;
	MemoStatistics->Visible=false;
	ButtonStart->Visible=false;
	ButtonDeleteTest->Visible=false;
	ButtonDelAll->Visible=false;
}
//---------------------------------------------------------------------------
void __fastcall TFStart::ProgramStart(){
	ButtonOfUsers->Visible = false;
	ButtonSettings->Visible = false;
	ButtonStartRandomTest->Visible = false;
	ButtonStartCreatedTest->Visible = false;
	TableOfRecordsButton->Visible = false;
	StartImage->Visible = false;

	WelkomeLabel->Top = (FStart->ClientHeight - ComboBoxUsers->Height - AddButton->Height - LabelHint->Height - 6*INDENTION)*5/12;
	ComboBoxUsers->Top = WelkomeLabel->Top + WelkomeLabel->Height + 2*INDENTION;
	LabelHint->Visible=true;
	LabelHint->Top = ComboBoxUsers->Top+ ComboBoxUsers->Height + 4*INDENTION;
	AddButton->Top = LabelHint->Top+LabelHint->Height+2*INDENTION;

	WelkomeLabel->Left = (FStart->ClientWidth - WelkomeLabel->Width)/2;
	ComboBoxUsers->Left = (FStart->ClientWidth - ComboBoxUsers->Width)/2;
	LabelHint->Width = ComboBoxUsers->Width;
	LabelHint->Left = (FStart->ClientWidth - LabelHint->Width)/2;
	AddButton->Left = (FStart->ClientWidth - AddButton->Width)/2;

	ComboBoxUsers->Text = "Выберите себя из списка...";
}

void __fastcall TFStart::Continue(){
	LabelHint->Visible = false;

	ComboBoxUsers->Top = StartImage->Top - 2*INDENTION - ComboBoxUsers->Height;
	WelkomeLabel->Top = ComboBoxUsers->Top - INDENTION - WelkomeLabel->Height;

	int dx = (StartImage->Width - 3*ButtonOfUsers->Width)/2;
	ButtonOfUsers->Left = StartImage->Left;
	ButtonSettings->Left = ButtonOfUsers->Left+ ButtonOfUsers->Width + dx;
	ButtonStartRandomTest->Left = ButtonSettings->Left + ButtonSettings->Width + dx;

	AddButton->Left = StartImage->Left;
	AddButton->Top = ButtonOfUsers->Top+ButtonOfUsers->Height+2*INDENTION;
	TableOfRecordsButton->Left = AddButton->Left + AddButton->Width +dx;
	ButtonStartCreatedTest->Left = TableOfRecordsButton->Left + TableOfRecordsButton->Width + dx;
	ButtonStartCreatedTest->Top = AddButton->Top;
	TableOfRecordsButton->Top = AddButton->Top;

	ButtonOfUsers->Visible = true;
	ButtonSettings->Visible = true;
	ButtonStartRandomTest->Visible = true;
	ButtonStartCreatedTest->Visible = true;
	TableOfRecordsButton->Visible = true;
	StartImage->Visible = true;
}

bool helper;
bool wasStart = false;

void __fastcall TFStart::FormActivate(TObject *Sender)
{
	if (wasStart == false) {
		ProgramStart();
	}
	AnsiString temp;
	Variant ExcelApplication,ExcelBooks,Sheet,Cell;
	wchar_t buffer[200];
	GetCurrentDirectory(sizeof(buffer),buffer);
	CURRENT_DIRECTORY=(AnsiString)buffer;

	if (wasChecked==true) {
		startChecking();
		wasChecked=false;
	}

	resolutionX=GetSystemMetrics(SM_CXSCREEN);
	resolutionY=GetSystemMetrics(SM_CYSCREEN);
	if (wasParametresChanged==false) {
		screenParametres();
		wasParametresChanged=true;
	}

	MemoStatistics->Height=IMAGE_WIDTH;
	TreeViewTestings->Height=IMAGE_HEIGHT-HEIGHT;
	MemoStatistics->Visible=false;
	ButtonStart->Visible=false;
	ButtonDeleteTest->Visible=false;
	ButtonDelAll->Visible=false;

	int amount=getBackgroundsAmount();
	if (amount!=0) {
		int numberOfBackGround=random(amount);
		AnsiString wayToBackGround=CURRENT_DIRECTORY+"\\backgrounds\\"+IntToStr(numberOfBackGround)+".jpg";
		try{
			StartImage->Picture->LoadFromFile(wayToBackGround);
		}
		catch(...){}
	}

	if (wasStretched==true) {
		ButtonStartCreatedTest->OnClick(Sender);
	}

	if (ComboBoxUsers->ItemIndex==-1){
		ButtonOfUsers->Enabled = false;
		ButtonSettings->Enabled = false;
		ButtonStartRandomTest->Enabled = false;
		ButtonStartCreatedTest->Enabled = false;
	}
	else{
		ButtonOfUsers->Enabled = true;
		ButtonSettings->Enabled = true;
		ButtonStartRandomTest->Enabled = true;
		ButtonStartCreatedTest->Enabled = true;
    }


	//ComboBoxUsers->Clear();
	 if (helper == false) { /// кто такой helper?????????????????????????????????????????????????


		int rowsCount,i;
		User *usersArray;

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
			FStart->Close();
		}

		usersArray=new User[rowsCount];     //дин. массив пользователей

		try{
			for (int i=1; i <= rowsCount; i++) {
				temp=Sheet.OlePropertyGet("Cells").OlePropertyGet("Item",i,1);//Текст клетки А1
				usersArray[i-1].setName(temp); //считывание имени пользователя с таблицы
			   /*	temp=Sheet.OlePropertyGet("Cells").OlePropertyGet("Item",i,5);
				usersArray[i-1].setSurname(temp); //считывание фамилии пользователя с таблицы
				usersurname = temp;
				temp=Sheet.OlePropertyGet("Cells").OlePropertyGet("Item",i,6);
				usersArray[i-1].setGroup(temp); //считывание группы пользователя с таблицы
				usergroup = temp; */
			}
			for (i = 0; i < rowsCount; i++) {
				ComboBoxUsers->Items->Add(usersArray[i].getName()/*+ " " + usersArray[i].getSurname() + " " + usersArray[i].getGroup()*/);
				//ComboBoxUsers->Items->Add(usersArray[i].getSurname());
				//ComboBoxUsers->Items->Add(usersArray[i].getGroup());

			}
		}
		catch(...){
			Application->Title="Ошибка";
			ShowMessage("Ошибка при считывании данных из файла\n"+CURRENT_DIRECTORY+"\\Users.xlsx");
			ExcelApplication.OleProcedure("Quit");

		}
		if (ComboBoxUsers->ItemIndex == -1) {
			ComboBoxUsers->Text = "Выберите себя из списка...";
		}


		//ComboBoxUsers->ItemIndex=0;//выбор первого пользователя по умолчанию

		//ComboBoxUsers->Text=Utf8ToAnsi("а б в");


		ExcelApplication.OleProcedure("Quit");
		delete []usersArray;
		helper = true;
	}

	else helper  = true;
}
//---------------------------------------------------------------------------

void __fastcall TFStart::ComboBoxUsersChange(TObject *Sender)
{
	ShowMessage("В данном поле регистрация невозможна.\nПожалуйста, пройдите регистрацию, нажав кнопку \"Добавить пользователя\"");
	ComboBoxUsers->ItemIndex=-1;
	ComboBoxUsers->Text = "Выберите себя из списка...";
	WelkomeLabel->Caption = "Добро пожаловать!";
	FStart->Activate();
}
//---------------------------------------------------------------------------


void __fastcall TFStart::AddButtonClick(TObject *Sender)
{
	FStart->Visible = false;
	FAddUser->Show();
	//NameBox->Visible = true;
	//OKButton->Visible = true;
}
//---------------------------------------------------------------------------

void addToCell(Variant Sheet,int row,int col,AnsiString value){
	Variant Cell;
	Cell=Sheet.OlePropertyGet("Cells").OlePropertyGet("Item",row,col);
	Cell.OlePropertySet("Value",StringToOleStr(value));
}

void __fastcall TFStart::OKButtonClick(TObject *Sender)
{
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
		ComboBoxUsers->Items->Add(Name);

		addToCell(Sheet,rowsCount+1,1,Name);

		NameBox->Visible = false;
		OKButton->Visible = false;
		ExcelApplication.OlePropertyGet("Workbooks").OlePropertyGet("Item",1).OleProcedure("Save");
	}
	catch(...){
		Application->Title="Ошибка";
		ShowMessage("Ошибка при записи данных в файл\n"+CURRENT_DIRECTORY+"\\Users.xlsx");
		ExcelApplication.OleProcedure("Quit");
		return;
	}

	ExcelApplication.OleProcedure("Quit");
}
//---------------------------------------------------------------------------

void __fastcall TFStart::NameBoxClick(TObject *Sender)
{
	NameBox->Clear();
}
//---------------------------------------------------------------------------


void __fastcall TFStart::ButtonOfUsersClick(TObject *Sender)
{
	User user;

	if (ComboBoxUsers->ItemIndex == -1) {
		Application->Title="Ошибка";
		ShowMessage("Выберите пользователя!");
		return;
	}

	UsersF->Show();
}
//---------------------------------------------------------------------------

void __fastcall TFStart::TableOfRecordsButtonClick(TObject *Sender)
{
	FStart->Visible = false;
	FormRecords->Show();
}
//---------------------------------------------------------------------------

TestInfo getInfoFromFile(FILE* file){
	TestInfo info;
	Application->Title="Ошибка";
	if (fread(&info.amountOfPassings,sizeof(int),1,file)!=1) ShowMessage("Ошибка при чтении числа прохождений");
	if (fread(&info.lastResult,sizeof(int),1,file)!=1) ShowMessage("Ошибка при чтении результата");
	if (fread(&info.lastTime,sizeof(int),1,file)!=1) ShowMessage("Ошибка при чтении времени");
	if (fread(&info.timeIndex,sizeof(int),1,file)!=1) ShowMessage("Ошибка при чтении временного индекса");
	if (fread(&info.questionsAmount,sizeof(int),1,file)!=1) ShowMessage("Ошибка при чтении числа вопросов");
	if (fread(&info.variantsAmount,sizeof(int),1,file)!=1) ShowMessage("Ошибка при чтении числа вариантов");

	if (fread(&info.isTimeLimited,sizeof(bool),1,file)!=1) ShowMessage("Ошибка при чтении временного ограничения");
	if (fread(&info.areImages,sizeof(bool),1,file)!=1) ShowMessage("Ошибка при чтении ограничения на иображения");

	fgets(info.user,20,file);
	fgets(info.theme,40,file);
	return info;
}

bool checkInfo(TestInfo* info){
	bool isRight=true;
	if (info->amountOfPassings >=1000 || info->amountOfPassings <1){
		info->amountOfPassings=1;
		isRight = false;
	}
	if (info->lastResult <0 || info->lastResult >10) {
		info->lastResult = 0;
		isRight = false;
	}
	if (info->lastTime < 0 || info->lastTime > 100000) {
		info->lastTime = 0;
		isRight = false;
	}
	if (info->timeIndex <0 || info->timeIndex > 2) {
		info->timeIndex = 2;
		isRight = false;
	}
	if (info->questionsAmount < 1 || info->questionsAmount > 80	) {
		info->questionsAmount = 1;
		isRight  =false;
	}
	if (info->variantsAmount < 4 || info->variantsAmount > 10) {
		info->variantsAmount = 4;
		isRight = false;
	}
	return isRight;
}

void __fastcall TFStart::TreeViewTestingsClick(TObject *Sender)
{
	ButtonStart->Enabled = true;
	try{
		if (wasNodeAdded==true) {
			AnsiString direction=CURRENT_DIRECTORY + "\\Testings\\" + TreeViewTestings->Selected->Text+".txt";

			try{
				TestInfo info;
				MemoStatistics->Lines->Clear();
				FILE* statisticsFile=fopen(direction.c_str(),"r");
				if (statisticsFile==NULL) {
					Application->Title="Ошибка";
					ShowMessage("Файл со статистикой тестирования был утерян, поэтому выбранное тестирование невозможно открыть.\nвыбранное тестирование будет удалено.");
					ButtonDeleteTest->Click();
					throw 1;
				}
				info = getInfoFromFile(statisticsFile);
				fclose(statisticsFile);

				if (checkInfo(&info) == false) {
					MemoStatistics->Lines->Add("Внимание! Данные о тестировании были повреждены, поэтому невозможно пройти выбранное тестирование.");
					MemoStatistics->Lines->Add("");
					ButtonStart->Enabled = false;
				}

				ButtonStart->Visible=true;
				ButtonDeleteTest->Visible=true;
				ButtonDelAll->Visible=true;
				MemoStatistics->Visible=true;
				AnsiString name=(AnsiString)info.user;
				AnsiString theme=(AnsiString)info.theme;
				AnsiString description="Вопросов: "+IntToStr(info.questionsAmount)+". Вариантов ответа: "+IntToStr(info.variantsAmount)+".";
				MemoStatistics->Lines->Add("Тема: "+theme);
				MemoStatistics->Lines->Add(description);
				if (info.isTimeLimited==true) {
					MemoStatistics->Lines->Add("Ограничение: "+FCharacteristics->RadioGroupTime->Items->Strings[info.timeIndex]+" на вопрос.");
				}
				else MemoStatistics->Lines->Add("Без временных ограничений");
				MemoStatistics->Lines->Add("");
				MemoStatistics->Lines->Add("\nСтатистика теста:");
				MemoStatistics->Lines->Add("Пройден: "+IntToStr(info.amountOfPassings)+" раз(а)");
				MemoStatistics->Lines->Add("");
				MemoStatistics->Lines->Add("Рекорд у пользователя "+name+":");
				MemoStatistics->Lines->Add("Результат: "+IntToStr(info.lastResult)+" балл(ов)");
				MemoStatistics->Lines->Add("Время: "+timeToString(info.lastTime));
				MemoStatistics->Lines->Delete(MemoStatistics->Lines->Count);
				MemoStatistics->Perform(EM_SCROLL,SB_LINEUP,0);
				areImagesInCreated=info.areImages;
			}
			catch(...){}
		}
	 }
	 catch(...){}
}
//---------------------------------------------------------------------------

AnsiString WayToCreatedTest;

void __fastcall TFStart::ButtonStartClick(TObject *Sender)
{
	WayToCreatedTest=CURRENT_DIRECTORY+"\\Testings\\"+TreeViewTestings->Selected->Text+".xlsx";
	AnsiString WayToCreatedTestStats=CURRENT_DIRECTORY+"\\Testings\\"+TreeViewTestings->Selected->Text+".txt";
	try{
		FILE* file=fopen(WayToCreatedTestStats.c_str(),"r");
		if (file == NULL) {
			Application->Title="Ошибка";
			ShowMessage("Ошибка при попытке запуска тестирования. Выберите другое тестирование.");
			return;
		}
		TestInfo info = getInfoFromFile(file);
		//file->Read(&info,sizeof(TestInfo));
		fclose(file);
		FCharacteristics->CheckBoxTime->Checked=info.isTimeLimited;
		FCharacteristics->RadioGroupTime->ItemIndex=info.timeIndex;
		FCharacteristics->CheckBoxTime->OnClick(Sender);
		wasSettingsChanged=true;
	}
	catch(...){}
	isRandomTest=false;
	open();
}
//---------------------------------------------------------------------------
void __fastcall TFStart::FormCreate(TObject *Sender)
{
	randomize();
}
//---------------------------------------------------------------------------
void __fastcall TFStart::ButtonDeleteTestClick(TObject *Sender)
{
	AnsiString wayToFile=CURRENT_DIRECTORY+"\\Testings\\"+TreeViewTestings->Selected->Text+".xlsx";
	DeleteFile(wayToFile.c_str());

	int i;
	for (i = 0; i < 5; i++) {
		wayToFile.Delete(wayToFile.Length(),1);
	}
	wayToFile+=".txt";
	DeleteFile(wayToFile.c_str());
	MemoStatistics->Lines->Clear();
	MemoStatistics->Visible=false;
	ButtonStart->Visible=false;
	ButtonDeleteTest->Visible=false;
	ButtonDelAll->Visible=false;
	ButtonStartCreatedTest->OnClick(Sender);
}
//---------------------------------------------------------------------------
void __fastcall TFStart::TimerForHidingTimer(TObject *Sender)
{
	static int i=0;
	i+=10;
	FStart->ClientWidth-=10;
	BevelFrame->Width-=10;
	if (abs(i-ADDITION_WIDTH)<=10) {
		int residue=abs(i-ADDITION_WIDTH);
		for (int j=0; j < residue; j++) {
			FStart->ClientWidth--;
			//BevelFrame->Width--;
		}
		i=0;
		ButtonStartCreatedTest->Enabled=true;
		LabelTestings->Enabled=true;
		TimerForHiding->Enabled=false;
	}
}
//---------------------------------------------------------------------------
void __fastcall TFStart::TimerForOpeningTimer(TObject *Sender)
{
	static int i=0;
	i+=10;
	FStart->ClientWidth+=10;
	BevelFrame->Width+=10;
	if (abs(i-ADDITION_WIDTH)<=10) {
		int residue=abs(i-ADDITION_WIDTH);
		for (int j=0; j < residue; j++) {
			FStart->ClientWidth++;
		}
		i=0;
		BevelFrame->Width=FStart->ClientWidth-BevelFrame->Left-INDENTION;
		ButtonStartCreatedTest->Enabled=true;
		LabelTestings->Visible=true;
		TimerForOpening->Enabled=false;
	}
}
//---------------------------------------------------------------------------
void __fastcall TFStart::LabelTestingsClick(TObject *Sender)
{
	if (LabelTestings->Caption=="Нет сохраненных тестирований"){
		wasStretched=false;
		ButtonStartCreatedTest->Enabled=false;
		LabelTestings->Visible=false;
		TimerForHiding->Enabled=true;
	}
}
//---------------------------------------------------------------------------
void __fastcall TFStart::ButtonDelAllClick(TObject *Sender)
{
	TTreeNode* node=TreeViewTestings->Items->GetFirstNode();
	while (node!=NULL){
		AnsiString file=CURRENT_DIRECTORY+"\\Testings\\"+node->Text+".xlsx";
		DeleteFile(file.c_str());
		AnsiString txtfile=CURRENT_DIRECTORY+"\\Testings\\"+node->Text+".txt";
		DeleteFile(txtfile.c_str());
		node=node->GetNext();
	}
	ButtonStartCreatedTest->Click();
}
//---------------------------------------------------------------------------
void startChecking(){
	AnsiString Documents= CURRENT_DIRECTORY+"\\Documents";
	if (DirectoryExists(Documents)==false){
		mkdir(Documents.c_str());
	}

	AnsiString Images=CURRENT_DIRECTORY+"\\Images";
	if (DirectoryExists(Images)==true) {
		chdir(Images.c_str());
		struct _finddata_t fileNames;
		intptr_t file;
		if ((file=_findfirst("*.jpg",&fileNames))!=-1) areImagesInFolder=true;
		else areImagesInFolder=false;
		chdir(CURRENT_DIRECTORY.c_str());
	}
	else areImagesInFolder=false;
	if (areImagesInFolder==false){
		Application->Title="Ошибка";
		ShowMessage("Изображения для тестирований были удалены. Требуется восстановить изображения в папке "+Images+" для корректного проведения тестирования.");
		mkdir(Images.c_str());
	}

	AnsiString media=CURRENT_DIRECTORY+"\\media";
	if (DirectoryExists(media)==false) {
		areMediaInFolder=false;
		Application->Title="Ошибка";
		ShowMessage("Мультимедийные файлы для видео- и аудиовопросов были удалены, данные типы вопросов не будут использоваться при тестировании.\nПри необходимости использования данных вопросов восстановите их в папке "+media);
	}
	else areMediaInFolder=true;

	bool areQuestions;
	AnsiString Questions=CURRENT_DIRECTORY+"\\Questions";
	if (DirectoryExists(Questions)==true) {
		chdir(Questions.c_str());
		struct _finddata_t fileNames;
		intptr_t file;
		if ((file=_findfirst("*.xlsx",&fileNames))!=-1) areQuestions=true;
		else areQuestions=false;
		chdir(CURRENT_DIRECTORY.c_str());

		FCharacteristics->ComboBoxTheme->Items->Clear();
		AnsiString fileName=Questions+"\\"+"Questions"+".xlsx";
		if (FileExists(fileName)==true) FCharacteristics->ComboBoxTheme->Items->Add("Все");

		fileName=Questions+"\\"+"Belarus in the Middle Age"+".xlsx";
		if (FileExists(fileName)==true) FCharacteristics->ComboBoxTheme->Items->Add("Древний мир и Средневековье");

		fileName=Questions+"\\"+"Belarus as a part of Grand Duchy of Lithuania"+".xlsx";
		if (FileExists(fileName)==true) FCharacteristics->ComboBoxTheme->Items->Add("Беларусь в составе ВКЛ");

		fileName=Questions+"\\"+"Belarus as a part of Poland"+".xlsx";
		if (FileExists(fileName)==true) FCharacteristics->ComboBoxTheme->Items->Add("Беларусь в составе РП");

		fileName=Questions+"\\"+"Belarus as a part of Russian empire"+".xlsx";
		if (FileExists(fileName)==true) FCharacteristics->ComboBoxTheme->Items->Add("Беларусь в составе Российской империи");

		fileName=Questions+"\\"+"Belarus in the beginning of XX century"+".xlsx";
		if (FileExists(fileName)==true) FCharacteristics->ComboBoxTheme->Items->Add("Беларусь в начале XX в.");

		fileName=Questions+"\\"+"Belarus in inter-war period"+".xlsx";
		if (FileExists(fileName)==true) FCharacteristics->ComboBoxTheme->Items->Add("Беларусь в межвоенный период");

		fileName=Questions+"\\"+"Belarus in WW2"+".xlsx";
		if (FileExists(fileName)==true) FCharacteristics->ComboBoxTheme->Items->Add("Беларусь в ВОВ");

		fileName=Questions+"\\"+"Belarus at a postwar time"+".xlsx";
		if (FileExists(fileName)==true) FCharacteristics->ComboBoxTheme->Items->Add("Послевоенное время");

		fileName=Questions+"\\"+"Media questions"+".xlsx";
		if (FileExists(fileName)==true) FCharacteristics->ComboBoxTheme->Items->Add("Мультимедийные вопросы");
	}
	else areQuestions=false;
	if (areQuestions==false) {
		Application->Title="Ошибка";
		ShowMessage("Невозможно пройти случайное тестирование, т.к. файлы с вопросами были удалены. Восстановите их в папке "+Questions);
		FStart->ButtonStartRandomTest->Enabled=false;
		mkdir(Questions.c_str());
	}

	AnsiString resources=CURRENT_DIRECTORY+"\\resourses";
	if (DirectoryExists(resources)==false) {
		mkdir(resources.c_str());
	}

	AnsiString Testings=CURRENT_DIRECTORY+"\\Testings";
	if (DirectoryExists(resources)==false) {
		mkdir(Testings.c_str());
	}
}

void __fastcall TFStart::ComboBoxUsersSelect(TObject *Sender)
{
	WelkomeLabel->Caption = "Добро пожаловать, " + ComboBoxUsers->Text ;
	WelkomeLabel->Left =ComboBoxUsers->Left + (ComboBoxUsers->Width - WelkomeLabel->Width)/2;
	if (wasStart == false) {
		wasStart = true;
		Continue();
	}
	FStart->Activate();
}
//---------------------------------------------------------------------------



