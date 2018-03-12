//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop

#include "UnitQuestion.h"
#include "FormStart.h"
#include "FormCharacteristics.h"
#include "FormTest.h"
#include "UnitTest.h"

#include <stdio.h>
#include <direct.h>
#include <io.h>
#include <ComObj.hpp>
#include <utilcls.h>
#include <JPEG.hpp>

//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma resource "*.dfm"
TFTest *FTest;
//---------------------------------------------------------------------------
__fastcall TFTest::TFTest(TComponent* Owner)
	: TForm(Owner)
{
}
//---------------------------------------------------------------------------

extern int LINE_LENGTH;
extern int IMAGE_WIDTH;
extern int IMAGE_HEIGHT;
extern int MAX_IMAGE_HEIGHT;
extern int MAX_IMAGE_WIDTH;
extern int WIDTH_INFO;
extern int WIDTH_ANSWERS;
extern int HEIGHT;
extern int FONT_HEIGHT;
extern int INDENTION;
extern int FORM_HEIGHT;
extern int FORM_WIDTH;

Question* TestQuestions;
extern int SECONDS_FOR_ANSWER;
const int FALSE_VARIANTS_AMOUNT=9;
const int IMAGES_AMOUNT=4;
const int ARRANGES=4;
const int MAX_VARIANTS_AMOUNT=10;
int globalCounter=-1;
Test testStatistics;
TImage* images[IMAGES_AMOUNT];
TPanel* panelsInfo[ARRANGES];
TPanel* panelsAnswers[ARRANGES];
TComboBox* comboBoxAnswers[ARRANGES];
TCheckBox* checkBoxes[MAX_VARIANTS_AMOUNT];
int variantsAmount;
int questionsAmount;
int upDownPosition=0;
int questionIndex=-1;
int help;
extern bool isRandomTest;
extern bool wasSettingsChanged;
extern bool areImagesInFolder;
extern bool areImagesInCreated;
bool areImages;
extern bool areMediaInFolder;
bool areMedia;
extern AnsiString username;
extern AnsiString CURRENT_DIRECTORY;
extern AnsiString WayToCreatedTest;
bool isMediaPlayerOpened=false;
bool wasAnyQuestionAnswered;
int seconds=0;
int secondsForAnswer=SECONDS_FOR_ANSWER;
bool isNeedToSave;
bool isSavedTestByProgram;
bool isCreatedTestByProgram;
bool isNeedToExit;
extern bool helper;
bool isAllowedToShow = false;

int generateRandomNumber(int* arrayOfUsedElements, int arraySize, int range){
	int i,number;
	do {
		number=random(range+1);
		for (i = 0; i < arraySize; i++) {
			if (number == arrayOfUsedElements[i]) {
				number=0;
				break;
			}
		}
	} while (number == 0);
	return number;
}
int generateRandomNumber(int range){
	int number=random(range+1);
	return number;
}

void zeroingArray(int* array,int size,int number){
	int i;
	for (i = 0; i < size; i++) {
		array[i]=number;
	}
}

void setVariants(Question* TestQuestions,int i, int* usedVariants, int question, Variant Sheet){
	int j,variant;
	AnsiString temp;
	for (j = 0; j < variantsAmount; j++){
		if (isRandomTest==true) {
			variant=generateRandomNumber(usedVariants, variantsAmount, FALSE_VARIANTS_AMOUNT);
		}
		else variant=j+1;
		usedVariants[j]=variant;
		variant+=2;   //смещение вправо, т.к. в А и В не то, что нужно
		temp=Sheet.OlePropertyGet("Cells").OlePropertyGet("Item",question,variant);//Текст клетки с вариантом ответа
		TestQuestions[i].setVariant(temp,j);
	}
}

void makingTest(){
	testStatistics.clearing();
	if (FCharacteristics->ComboBoxTheme->ItemIndex==-1) FCharacteristics->ComboBoxTheme->ItemIndex=0;

	int rowsCount,i,j;
	AnsiString temp,tempWay,WayToFile,WayToPicture;
	int* usedQuestions;
	int* usedVariants;
	Variant ExcelApplication,ExcelBooks,Sheet,Cell;

	if (isRandomTest==true){
		switch (FCharacteristics->ComboBoxTheme->ItemIndex) {
			case -1:
			case 0: WayToFile=CURRENT_DIRECTORY+"\\Questions\\Questions.xlsx"; help = 1; break;
			case 1: WayToFile=CURRENT_DIRECTORY+"\\Questions\\Belarus in the Middle Age.xlsx";  help = 2; break;
			case 2: WayToFile=CURRENT_DIRECTORY+"\\Questions\\Belarus as a part of Grand Duchy of Lithuania.xlsx";  help = 3;break;
			case 3: WayToFile=CURRENT_DIRECTORY+"\\Questions\\Belarus as a part of Poland.xlsx";  help = 4; break;
			case 4: WayToFile=CURRENT_DIRECTORY+"\\Questions\\Belarus as a part of Russian empire.xlsx"; help = 5; break;
			case 5: WayToFile=CURRENT_DIRECTORY+"\\Questions\\Belarus in the beginning of XX century.xlsx";  help = 6; break;
			case 6: WayToFile=CURRENT_DIRECTORY+"\\Questions\\Belarus in inter-war period.xlsx";  help = 7; break;
			case 7: WayToFile=CURRENT_DIRECTORY+"\\Questions\\Belarus in WW2.xlsx";  help = 8; break;
			case 8: WayToFile=CURRENT_DIRECTORY+"\\Questions\\Belarus at a postwar time.xlsx";  help = 9; break;
			case 9: WayToFile=CURRENT_DIRECTORY+"\\Questions\\Media questions";  help = 10; break;
		}
	}
	else {
		WayToFile=WayToCreatedTest;
	}

	ExcelApplication=CreateOleObject("Excel.Application");
	try{
		ExcelBooks=ExcelApplication.OlePropertyGet("Workbooks").OlePropertyGet("Open",StringToOleStr(WayToFile));
		Sheet=ExcelBooks.OlePropertyGet("Worksheets",1);
		rowsCount=Sheet.OlePropertyGet("UsedRange").OlePropertyGet("Rows").OlePropertyGet("Count");\
	}
	catch(...) {
		Application->Title="Ошибка";
		ShowMessage("Ошибка при открытии файла\n"+WayToFile+"\nПроверьте наличие файла в директории\n"+CURRENT_DIRECTORY+"\\Questions");
		ExcelApplication.OleProcedure("Quit");
		FStart->Close();
	}

	try{
		if (rowsCount < StrToInt(FCharacteristics->EditQuestionsAmount->Text)) {
			FCharacteristics->UpDown1->Position=rowsCount;
			FCharacteristics->EditQuestionsAmount->Text=IntToStr(rowsCount);
		}

		questionsAmount=FCharacteristics->EditQuestionsAmount->Text.ToInt();
		variantsAmount=FCharacteristics->EditVariantsAmount->Text.ToInt()-1;

		if (isRandomTest==false) {
			questionsAmount=StrToInt(Sheet.OlePropertyGet("Cells").OlePropertyGet("Item",1,18));
			variantsAmount= StrToInt(Sheet.OlePropertyGet("Cells").OlePropertyGet("Item",1,19))-1;
		}
	}
	catch(...){
		Application->Title="Ошибка";
		ShowMessage("Выбрано неверное количество вариантов либо вопросов");
		ExcelApplication.OleProcedure("Quit");
		isNeedToExit=true;
		FTest->Close();
	}

	try{
		TestQuestions=new Question[questionsAmount];
		usedQuestions=new int[questionsAmount];
		usedVariants=new int[variantsAmount];

		for (i = 0; i < IMAGES_AMOUNT; i++) {
			images[i]=new TImage(FTest);
		}
		for (i = 0; i < ARRANGES; i++) {
			panelsInfo[i]=new TPanel(FTest);
			panelsAnswers[i]=new TPanel(FTest);
			comboBoxAnswers[i]=new TComboBox(FTest);
			comboBoxAnswers[i]->MaxLength=100;
		}
		for (i = 0; i < MAX_VARIANTS_AMOUNT; i++) {
			checkBoxes[i]=new TCheckBox(FTest);
		}

		for (i = 0; i < IMAGES_AMOUNT; i++) {
			images[i]->Visible=false;
		}
		for (i = 0; i < ARRANGES; i++) {
			panelsInfo[i]->Visible=false;
			panelsAnswers[i]->Visible=false;
		}

		zeroingArray(usedQuestions,questionsAmount,0);
		//zeroingArray(usedVariants,variantsAmount,0);

		for (i = 0; i < questionsAmount; i++) {
			zeroingArray(usedVariants,variantsAmount,0);
			int j,question;
			if (isRandomTest==true) {
				question=generateRandomNumber(usedQuestions, questionsAmount, rowsCount);
			}
			else question=i+1;
			usedQuestions[i]=question;

			temp=Sheet.OlePropertyGet("Cells").OlePropertyGet("Item",question,20);//Текст клетки T"question" с типом вопроса
			if (areMedia==false) {
				if (temp=="music" || temp=="video") {
					i--;
					continue;
				}
			}
			if (areImages==false) {
				if (temp=="arrangepictures") {
					i--;
					continue;
				}
			}
			TestQuestions[i].setQuestionType(temp);

			temp=Sheet.OlePropertyGet("Cells").OlePropertyGet("Item",question,23);//Текст клетки W"question" со ссылкой
			TestQuestions[i].setReference(temp);

			temp=Sheet.OlePropertyGet("Cells").OlePropertyGet("Item",question,1);//Текст клетки А"question" с текстом вопроса
			TestQuestions[i].setText(temp);

			if (areImages==true || (isRandomTest==false && areImagesInCreated==true)) {
			}
			else{
				if (TestQuestions[i].getQuestionType()=="choose" || TestQuestions[i].getQuestionType()=="write") {
					temp=Sheet.OlePropertyGet("Cells").OlePropertyGet("Item",question,12);//Текст клетки L"question" с доп. формулировкой
					TestQuestions[i].setText(temp);
				}
            }

			if (TestQuestions[i].getQuestionType()=="choose" || TestQuestions[i].getQuestionType()=="write" || TestQuestions[i].getQuestionType()=="music" || TestQuestions[i].getQuestionType()=="video") {
				temp=Sheet.OlePropertyGet("Cells").OlePropertyGet("Item",question,2);//Текст клетки B"question"  с правильным ответом
				TestQuestions[i].setRightVariant(temp);
			}
			if (TestQuestions[i].getQuestionType()=="somevariants" || TestQuestions[i].getQuestionType()=="arrangepictures") {
				temp=Sheet.OlePropertyGet("Cells").OlePropertyGet("Item",question,2);
				TestQuestions[i].setRightVariantsAmount(StrToInt(temp));
			}

			temp=Sheet.OlePropertyGet("Cells").OlePropertyGet("Item",question,13);//Текст клетки M"question" со сложностью вопроса
			TestQuestions[i].setDifficulty(temp.ToInt());
			testStatistics.addMaxPoints(temp.ToInt());

			TestQuestions[i].setIsAnswered(false);

			for (j=0; j<IMAGES_AMOUNT; j++) {
				temp=Sheet.OlePropertyGet("Cells").OlePropertyGet("Item",question,j+14);//Текст клеток N O P R "question"  с именем изображений
				if (temp!="-") {
					if (isRandomTest==true) TestQuestions[i].setPicture(CURRENT_DIRECTORY+"\\Images\\"+temp,j);
					else  TestQuestions[i].setPicture(temp,j);
				}
				else TestQuestions[i].setPicture(temp,j);                //choose-выбор вариантов ответа
			}                                                            //write-вписать ответ
			if (TestQuestions[i].getQuestionType()=="choose") {          //arrange - расставить
				setVariants(TestQuestions,i,usedVariants,question,Sheet);     //somevariants - несколько прав. отв.
			}                                                                 //arrangepictures - расстанвить по картинкам
			if (TestQuestions[i].getQuestionType()=="arrange") {              //music
				for (j=0; j < ARRANGES*2; j++) {                              //video
					temp=Sheet.OlePropertyGet("Cells").OlePropertyGet("Item",question,j+2);
					TestQuestions[i].setVariant(temp,j);
				}
			}
			if (TestQuestions[i].getQuestionType()=="somevariants") {
				for (j=0; j <= variantsAmount+1; j++) {
					temp=Sheet.OlePropertyGet("Cells").OlePropertyGet("Item",question,j+3);
					TestQuestions[i].setVariant(temp,j);
				}
			}
			if (TestQuestions[i].getQuestionType()=="arrangepictures") {
				for (j=0; j < TestQuestions[i].getRightVariantsAmount(); j++) {
					temp=Sheet.OlePropertyGet("Cells").OlePropertyGet("Item",question,j+3);
					TestQuestions[i].setVariant(temp,j);
				}
			}
			if (TestQuestions[i].getQuestionType()=="music" || TestQuestions[i].getQuestionType()=="video") {
				setVariants(TestQuestions,i,usedVariants,question,Sheet);
				temp=Sheet.OlePropertyGet("Cells").OlePropertyGet("Item",question,21);   //U"question"
				if (isRandomTest==true) TestQuestions[i].setWayToMediaFile(CURRENT_DIRECTORY+"\\media\\"+temp);
				else TestQuestions[i].setWayToMediaFile(temp);
			}
			if (TestQuestions[i].getQuestionType()=="video") {
				temp=Sheet.OlePropertyGet("Cells").OlePropertyGet("Item",question,24);   //X"question"
				TestQuestions[i].setVideoResolutionX(StrToInt(temp));
				temp=Sheet.OlePropertyGet("Cells").OlePropertyGet("Item",question,25);   //Y"question"
				TestQuestions[i].setVideoResolutionY(StrToInt(temp));
			}
		}
		ExcelApplication.OleProcedure("Quit");
	}

	catch(...){
		Application->Title="Ошибка";
		ShowMessage("Не удается создать тестирование.\nВозможно, файл\n"+WayToFile+"\nбыл поврежден");
		ExcelApplication.OleProcedure("Quit");
		delete []usedVariants;
		delete []usedQuestions;
		FTest->Close();
	}
}

void __fastcall TFTest::FormActivate(TObject *Sender)
{
	TimerTime->Enabled=true;
	globalCounter=0;
	FTest->ClientHeight=FORM_HEIGHT;
	FTest->ClientWidth=FORM_WIDTH;
	FTest->Top=0;
	Bevel1->Width=FTest->ClientWidth;
	Bevel1->Height=FTest->ClientHeight;
	Bevel1->Top=0;
	Bevel1->Left=0;
	isNeedToSave=true;

	if (FCharacteristics->RadioGroupArrowsOrDropDowns->ItemIndex==0) {
		FTest->MenuArrow->Checked=true;
	}
	else if (FCharacteristics->RadioGroupArrowsOrDropDowns->ItemIndex==1) {
		FTest->MenuDropDown->Checked=true;
	}
	if (areImagesInFolder==true) {
		areImages=FCharacteristics->CheckBoxImages->Checked;
	}
	else areImages=false;
	if (areMediaInFolder==true) {
		areMedia=FCharacteristics->CheckBoxMedia->Checked;
	}
	else areMedia=false;
	if (isRandomTest==false && areImagesInCreated==true) {
		areImages=true;
	}

	PanelTime->Font->Color=clBlack;
	ButtonSaveTest->Enabled=true;
	EditTestName->Enabled=true;
	ButtonCreateDocument->Enabled=true;

	isNeedToExit=false;
	isAllowedToShow = false;
	isSavedTestByProgram=false;
	isCreatedTestByProgram=false;
	wasAnyQuestionAnswered=false;
	ProgressBar1->Visible=false;
	MemoCheck->Visible=false;
	LabelQuestionNumber->Visible=false;
	LabelSaving->Visible=false;
	LabelEnd->Visible=false;;
	LabelResult->Visible=false;
	ButtonGoBack->Visible=false;
	MemoQuestionText->Visible=false;
	RadioGroupVariants->Visible=false;
	ButtonAnswer->Visible=false;
	ButtonStopTest->Visible=false;
	ButtonCreateDocument->Visible=false;
	ButtonSaveTest->Visible=false;
	EditTestName->Visible=false;
	LabelMark->Visible=false;
	LabelEnterAnswer->Visible=false;
	StringGridResults->Visible=false;
	ScrollBar1->Visible=false;
	MediaPlayer1->Visible=false;
	PanelVideo->Visible=false;
	ButtonCreateDocument->Caption="Оформить отчет";

	makingTest();

	seconds=0;
	if (FCharacteristics->CheckBoxTime->Checked==true){
		secondsForAnswer=SECONDS_FOR_ANSWER;
		PanelTime->Caption=timeToString(secondsForAnswer);
	}
	else{
		PanelTime->Caption=timeToString(seconds);
		TimerTime->Enabled=true;
	}

	EditAnswer->Width=IMAGE_WIDTH;
	ButtonPreviousQuestion->Left=INDENTION;
	ButtonPreviousQuestion->Top=FTest->ClientHeight-ButtonPreviousQuestion->Height-INDENTION;

	ButtonNextQuestion->Left=FTest->ClientWidth-ButtonNextQuestion->Width-INDENTION;
	ButtonNextQuestion->Top=ButtonPreviousQuestion->Top;

	MemoQuestionText->Left=2*INDENTION;
	MemoQuestionText->Width=FTest->ClientWidth-4*INDENTION;
	MemoQuestionText->Top=6*INDENTION;
	MemoQuestionText->Font->Height=FONT_HEIGHT*1.4;

	ProgressBar1->Left=2*INDENTION;
	ProgressBar1->Width=FTest->ClientWidth-8*INDENTION-ButtonStopTest->Width-PanelTime->Width;
	ProgressBar1->Max=(MemoQuestionText->Width/questionsAmount)*questionsAmount;
	ProgressBar1->Step=ProgressBar1->Max/questionsAmount;
	ProgressBar1->Position=0;

	ButtonStopTest->Left=ProgressBar1->Left+ProgressBar1->Width+2*INDENTION;
	ButtonStopTest->Top=ProgressBar1->Top - (ButtonStopTest->Height - ProgressBar1->Height)/2;

	PanelTime->Left=ButtonStopTest->Left+ButtonStopTest->Width+2*INDENTION;
	PanelTime->Top=ButtonStopTest->Top;

	ButtonAnswer->Left=(FTest->ClientWidth-ButtonAnswer->Width)/2;

	LabelBottomLine->Font->Height=FONT_HEIGHT*1.5;
	LabelBottomLine->Caption="Вопрос "+IntToStr(globalCounter+1)+" из "+IntToStr(questionsAmount);
	LabelBottomLine->Left=ButtonPreviousQuestion->Left + ButtonPreviousQuestion->Width + (FTest->ClientWidth - ButtonPreviousQuestion->Width - ButtonNextQuestion->Width - LabelBottomLine->Width - 4*INDENTION)/2;
	LabelBottomLine->Top=ButtonNextQuestion->Top + (ButtonNextQuestion->Height - LabelBottomLine->Height)/2;

	ButtonAnswer->Visible=true;
	ProgressBar1->Visible=true;
	if (FCharacteristics->CheckBoxTime->Checked==false) {
		ButtonPreviousQuestion->Visible=true;
		ButtonNextQuestion->Visible=true;
	}
	else {
		ButtonPreviousQuestion->Visible=false;
		ButtonNextQuestion->Visible=false;
    }
	LabelBottomLine->Visible=true;
	ButtonStopTest->Visible=true;
	PanelTime->Visible=true;
	FTest->ProgressBar1->StepIt();

	changeQuestion(globalCounter);
}
//---------------------------------------------------------------------------
void __fastcall TFTest::FormClose(TObject *Sender, TCloseAction &Action)
{
	if (isNeedToExit==false) {
		stopTest();
		FTest->Visible=false;
		FStart->Visible=true;
		deleteTest();
		FStart->Show();
		isNeedToExit=false;
	}
	else {
        FTest->Visible=false;
		FStart->Visible=true;
	}
}
//---------------------------------------------------------------------------
void __fastcall TFTest::FormCreate(TObject *Sender)
{
	randomize();
}
//---------------------------------------------------------------------------
bool changeToNotAnsweredQuestion(){
	int i;
	bool isNotAnsweredQuestion=false;

	for (i = globalCounter+1; i < questionsAmount; i++) {
		if (TestQuestions[i].getIsAnswered()==false) {
			isNotAnsweredQuestion=true;
			globalCounter=i;
			changeQuestion(globalCounter);
			break;
		}
	}

	if (isNotAnsweredQuestion==false) {
		for (i = 0; i <= globalCounter; i++) {
			if (TestQuestions[i].getIsAnswered()==false) {
				isNotAnsweredQuestion=true;
				globalCounter=i;
				changeQuestion(globalCounter);
				break;
			}
		}
	}

	return isNotAnsweredQuestion;
}

void writeInfoToFile(struct TestInfo info, FILE* file){
	Application->Title="Ошибка";
	if (fwrite(&info.amountOfPassings,sizeof(int),1,file)!=1) ShowMessage("Ошибка при записи числа прохождений");
	if (fwrite(&info.lastResult,sizeof(int),1,file)!=1) ShowMessage("Ошибка при записи результата");
	if (fwrite(&info.lastTime,sizeof(int),1,file)!=1) ShowMessage("Ошибка при записи времени");
	if (fwrite(&info.timeIndex,sizeof(int),1,file)!=1) ShowMessage("Ошибка при записи временного индекса");
	if (fwrite(&info.questionsAmount,sizeof(int),1,file)!=1) ShowMessage("Ошибка при записи числа вопросов");
	if (fwrite(&info.variantsAmount,sizeof(int),1,file)!=1) ShowMessage("Ошибка при записи числа вариантов");

	if (fwrite(&info.isTimeLimited,sizeof(bool),1,file)!=1) ShowMessage("Ошибка при записи временного ограничения");
	if (fwrite(&info.areImages,sizeof(bool),1,file)!=1) ShowMessage("Ошибка при записи ограничения на изображения");

	fputs(info.user,file);
	fputs(info.theme,file);
}

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

AnsiString convertUnit(int value, AnsiString str){
	AnsiString result="";
	switch (value%10) {
		case 0: result="";  break;
		case 1:
			if (value!=11) result=IntToStr(value)+" "+str+"у";
			else result=IntToStr(value)+" "+str;
			break;
		case 2:
		case 3:
		case 4:
			if (value!=14 && value!=12 && value!=13) result=IntToStr(value)+" "+str+"ы";
			else result=IntToStr(value)+" "+str;
			break;
		default: result=IntToStr(value)+" "+str; break;
	}
	return result;
}

AnsiString convertTime(){
	int minute=testStatistics.getTime()/60;
	int second=testStatistics.getTime() - (testStatistics.getTime()/60)*60;
	AnsiString result=convertUnit(minute,"минут");
	if (result!="") {
		result+=", ";
	}
	result+=convertUnit(second,"секунд");
	if (result=="") return "0 секунд";
	else return result;
}

void __fastcall TFTest::saveResults(){
	AnsiString WayToFile;      int i;
			AnsiString username,cellb, cellc,celld,cellh,celli, pasLastTest, pasAllTests,pasAverageScore, pasAllTime;
			Variant ExcelApplication,ExcelBooks,Sheet,Cell;
			WayToFile=CURRENT_DIRECTORY+"\\Users.xlsx";
			int position = FStart->ComboBoxUsers->ItemIndex+1;
			const int COLS_AMOUNT = 36;
			int rowsCount;
			int pasTen = 10, pasTwenty = 20, pasThirty = 30, pasForty = 40, pasFifty = 50;

			try{
				ExcelApplication=CreateOleObject("Excel.Application");
				ExcelBooks=ExcelApplication.OlePropertyGet("Workbooks").OlePropertyGet("Open",StringToOleStr(WayToFile));
				Sheet=ExcelBooks.OlePropertyGet("Worksheets",1);
				rowsCount=Sheet.OlePropertyGet("UsedRange").OlePropertyGet("Rows").OlePropertyGet("Count");
			}
			catch(...){
				ShowMessage("Ошибка при обращении к файлу\n"+WayToFile);
				ExcelApplication.OleProcedure("Quit");
				return;
			}
   ///////////////////////////////////////////////////////////////////////////////////////
			try{
				for (i = 0; i < rowsCount; i++) {
					cellb=Sheet.OlePropertyGet("Cells").OlePropertyGet("Item",i+1,1);//Текст клетки A

					if (FStart->ComboBoxUsers->Text==cellb) {
						cellc=Sheet.OlePropertyGet("Cells").OlePropertyGet("Item",i+1,3);//Текст клетки C
						celld=Sheet.OlePropertyGet("Cells").OlePropertyGet("Item",i+1,4);//Текст клетки D
						cellh=Sheet.OlePropertyGet("Cells").OlePropertyGet("Item",i+1,8);//Текст клетки H

						pasAllTests = Sheet.OlePropertyGet("Cells").OlePropertyGet("Item",i+1,pasTwenty + help);
						pasAverageScore = Sheet.OlePropertyGet("Cells").OlePropertyGet("Item",i+1,pasThirty + help);
						pasAllTime = Sheet.OlePropertyGet("Cells").OlePropertyGet("Item",i+1,pasFifty + help);

						celld = FloatToStr(StrToFloat(celld)*StrToInt(cellc));
						cellc = IntToStr(StrToInt(cellc)+1);
						celld = FloatToStr(StrToFloat(celld)+testStatistics.testResult());
						celld=FloatToStr((float)StrToFloat(celld)/StrToInt(cellc));

						addToCell(Sheet,i+1,2,IntToStr(testStatistics.testResult()));
						addToCell(Sheet,i+1,3,cellc);
						addToCell(Sheet,i+1,4,celld);

						cellh = FloatToStr(StrToFloat(cellh)+testStatistics.getTime());
						addToCell(Sheet,i+1,7,IntToStr(testStatistics.getTime()));
						addToCell(Sheet,i+1,8,cellh);

						//pasAllTests = Sheet.OlePropertyGet("Cells").OlePropertyGet("Item",i+1,pasTwenty + help);
						//pasAverageScore = Sheet.OlePropertyGet("Cells").OlePropertyGet("Item",i+1,pasThirty + help);


						pasAverageScore = FloatToStr(StrToFloat(pasAverageScore)*StrToFloat(pasAllTests));     //////
						pasAllTests = FloatToStr(StrToFloat(pasAllTests)+1);     ///////////////////////////////////////
						pasAverageScore = FloatToStr(StrToFloat(pasAverageScore)+testStatistics.testResult());
						pasAverageScore=FloatToStr((float)StrToFloat(pasAverageScore)/StrToFloat(pasAllTests));
						pasAllTime = FloatToStr(StrToFloat(pasAllTime)+testStatistics.getTime());

						addToCell(Sheet,i+1,pasTen + help,IntToStr(testStatistics.testResult()));
						addToCell(Sheet,i+1,pasTwenty + help,pasAllTests);
						addToCell(Sheet,i+1,pasThirty + help,pasAverageScore);
						addToCell(Sheet,i+1,pasForty+ help,IntToStr(testStatistics.getTime()));
						addToCell(Sheet,i+1,pasFifty + help,pasAllTime);

					}

				}

				ExcelApplication.OlePropertyGet("Workbooks").OlePropertyGet("Item",1).OleProcedure("Save");
				ExcelApplication.OleProcedure("Quit");
				helper = true;
				//FStart->ComboBoxUsers->ItemIndex=0;//выбор первого пользователя по умолчанию
			}

			catch(...){
				Application->Title="Ошибка";
				ShowMessage("Ошибка при сохранении результатов тестирования");
				ExcelApplication.OleProcedure("Quit");
			}
}


void __fastcall TFTest::stopTest(){
	testStatistics.setIsEndOfTest(true);
	MemoQuestionText->Visible=false;
	RadioGroupVariants->Visible=false;
	ButtonAnswer->Visible=false;
	ButtonStopTest->Visible=false;
	EditAnswer->Visible=false;
	LabelEnterAnswer->Visible=false;
	ButtonUp->Visible=false;
	ButtonDown->Visible=false;
	ButtonUpGeneral->Visible=false;
	ButtonDownGeneral->Visible=false;
	ButtonRight->Visible=false;
	ButtonLeft->Visible=false;
	ButtonRightGeneral->Visible=false;
	ButtonLeftGeneral->Visible=false;
	ProgressBar1->Visible=false;
	ButtonNextQuestion->Visible=false;
	ButtonPreviousQuestion->Visible=false;
	LabelBottomLine->Visible=false;
	PanelVideo->Visible=false;
	ScrollBar1->Visible=false;
	PanelTime->Visible=false;
	TimerTime->Enabled=false;
	testStatistics.setTime(seconds);
	seconds=0;
	PanelTime->Caption="";
	if (isMediaPlayerOpened==true) {
		MediaPlayer1->Close();
		isMediaPlayerOpened=false;
	}
	MediaPlayer1->Visible=false;
					  int i;
	for (i = 0; i < IMAGES_AMOUNT; i++) {
		images[i]->Visible=false;
	}
	for (i = 0; i < ARRANGES; i++) {
		panelsInfo[i]->Visible=false;
		panelsAnswers[i]->Visible=false;
		comboBoxAnswers[i]->Visible=false;
	}
	for (i = 0; i < variantsAmount+1; i++) {
		checkBoxes[i]->Visible=false;
	}

	AnsiString name=FStart->ComboBoxUsers->Text+" "+dateNow();
	name = findFiles(name,"\\Testings",".xlsx");
	EditTestName->Text=name;

	if (isNeedToSave==true) {
		LabelSaving->Top=(FTest->ClientHeight-LabelSaving->Height)/3;
		LabelSaving->Left=(FTest->ClientWidth-LabelSaving->Width)/2;
		LabelQuestionNumber->Left=(FTest->ClientWidth-100)/2;
		LabelQuestionNumber->Top=LabelSaving->Top+LabelSaving->Height+2*INDENTION;
		LabelSaving->Visible=true;
		ProgressBar1->Top=LabelSaving->Top+LabelSaving->Height+6*INDENTION;

		if (FCharacteristics->CheckBoxAutoSaveTest->Checked==true && FCharacteristics->CheckBoxAutoCreateDocument->Checked==true) {
			ButtonCreateDocument->Visible=false;
			ButtonGoBack->Left=(FTest->ClientWidth-ButtonGoBack->Width)/2;
		}
		if (FCharacteristics->CheckBoxAutoSaveTest->Checked==true && isRandomTest==true) {
			ButtonGoBack->Enabled=false;
			ButtonSaveTest->Enabled=false;
			EditTestName->Enabled=false;
			LabelSaving->Caption="Сохранение тестирования";
			LabelSaving->Left=(FTest->ClientWidth-LabelSaving->Width)/2;
			isSavedTestByProgram=true;
			ButtonSaveTest->Click();
			ButtonGoBack->Enabled=true;
		}
		if (FCharacteristics->CheckBoxAutoCreateDocument->Checked==true) {
			ButtonGoBack->Enabled=false;
			ButtonCreateDocument->Enabled=false;
			LabelSaving->Caption="Создание отчета";
			LabelSaving->Left=(FTest->ClientWidth-LabelSaving->Width)/2;
			isCreatedTestByProgram=true;
			ButtonCreateDocument->Click();
			ButtonGoBack->Enabled=true;
		}
		isNeedToSave=false;
		LabelSaving->Visible=false;
		ProgressBar1->Visible=false;
		ProgressBar1->Top=INDENTION;
	}

	int result=testStatistics.getCollectedPoints();
	int residue;
	AnsiString temp;
	if (result<11 || result>14) {    //11-14 очков, а не 1 очко, 2-4 очка
		residue=result%10;
	}
	else residue=5; // 5 очков
	switch (residue) {
		case 1: temp=" балл "; break;
		case 2:
		case 3:
		case 4: temp=" балла "; break;
		default: temp=" баллов "; break;
	}

	if (isRandomTest==false){
		TestInfo info;
		AnsiString wayToFile=WayToCreatedTest.Delete(WayToCreatedTest.Length(),1);
		for (i = 0; i < 4; i++) {
			wayToFile.Delete(wayToFile.Length(),1);
		}
		wayToFile+=".txt";

		FILE* file=fopen(wayToFile.c_str(),"r");
		if (file!=NULL) {
			//if (fread(&info,sizeof(TestInfo),1,file)==1) {
				info = getInfoFromFile(file);
				fclose(file);
				AnsiString name=FStart->ComboBoxUsers->Text+"\n";
				info.amountOfPassings+=1;
				if (info.lastResult == testStatistics.testResult()){
					if (info.lastTime > testStatistics.getTime()){
						info.lastTime=testStatistics.getTime();
						strcpy(info.user,name.c_str());
					}
				}
				else if (info.lastResult < testStatistics.testResult()){
					info.lastResult=testStatistics.testResult();
					info.lastTime=testStatistics.getTime();
					strcpy(info.user,name.c_str());
				}

				file=fopen(wayToFile.c_str(),"w");
				writeInfoToFile(info, file);
				//if (fwrite(&info,sizeof(TestInfo),1,file)!=1) ShowMessage("Ошибка при сохранении результатов теста");
			/*}
			else{
				ShowMessage("Ошибка при считывании результатов теста. Предыдущие результаты обнулены.");
				info.amountOfPassings=1;
				info.lastResult = testStatistics.testResult();
				info.lastTime = testStatistics.getTime();
				strcpy(info.user,name.c_str());
				strcpy(info.theme,"Не определена");
				info.isTimeLimited = FCharacteristics->CheckBoxTime->Checked;
				info.areImages = FCharacteristics->CheckBoxImages->Checked;
				info.timeIndex = FCharacteristics->RadioGroupTime->ItemIndex;
				info.questionsAmount = questionsAmount;
				info.variantsAmount = variantsAmount;
                file=fopen(wayToFile.c_str(),"w");
				if (fwrite(&info,sizeof(TestInfo),1,file)!=1) ShowMessage("Ошибка при повторном сохранении результатов теста");
			} */
		}
		else {
			Application->Title="Ошибка";
			ShowMessage("Невозможно найти файл с прошлыми результатами теста. Статистика будет обновлена");
			info.amountOfPassings=1;
			info.lastResult=testStatistics.testResult();
			info.lastTime=testStatistics.getTime();
			strcpy(info.user,(char*)FStart->ComboBoxUsers->Text.c_str());
			strcpy(info.theme,"Не определена");

			file=fopen(wayToFile.c_str(),"w");
			if (fwrite(&info,sizeof(TestInfo),1,file)!=1) ShowMessage("Ошибка при сохранении результатов теста");
		}
		if (file!=NULL) fclose(file);
	}

	AnsiString text="Вы набрали "+IntToStr(result)+temp+"из "+IntToStr(testStatistics.getMaxPoints());
	LabelResult->Caption=text;
	LabelMark->Caption="Оценка: "+IntToStr(testStatistics.testResult());

	LabelEnd->Caption="Тестирование пройдено за "+convertTime();
	LabelEnd->Left=(FTest->ClientWidth-(LabelEnd->Width+5*INDENTION))/2;
	LabelEnd->Top=FTest->ProgressBar1->Top+5*INDENTION;
	LabelResult->Left=(FTest->ClientWidth-(LabelResult->Width+5*INDENTION))/2;
	LabelResult->Top=LabelEnd->Top + LabelEnd->Height +2*INDENTION;
	LabelMark->Left=(FTest->ClientWidth-LabelMark->Width)/2;
	LabelMark->Top=LabelResult->Top+LabelResult->Height+2*INDENTION;

	StringGridResults->RowCount=questionsAmount+1;
	StringGridResults->Cells[0][0]="№ вопроса";
	StringGridResults->Cells[1][0]="Правильность";
	StringGridResults->FixedRows=1;
	StringGridResults->DefaultRowHeight=3*INDENTION;
	StringGridResults->DefaultColWidth=WIDTH_INFO*2/3;
	StringGridResults->Width=2*StringGridResults->DefaultColWidth+2*StringGridResults->GridLineWidth+5;
	if (questionsAmount>=8) {
		StringGridResults->Height=8*StringGridResults->DefaultRowHeight+9*StringGridResults->GridLineWidth+3;
		StringGridResults->ScrollBars=ssVertical;
	}
	else {
		StringGridResults->ScrollBars=ssNone;
		StringGridResults->Height=(questionsAmount+1)*StringGridResults->DefaultRowHeight+(questionsAmount+2)*StringGridResults->GridLineWidth+3;
	}

	StringGridResults->Left=(FTest->ClientWidth-StringGridResults->Width)/2;
	StringGridResults->Top=LabelMark->Top+LabelMark->Height+3*INDENTION;

	for (i = 0; i < questionsAmount; i++) {
		if (TestQuestions[i].getIsRightAnswer()==true) {
			StringGridResults->Cells[1][i+1]="Правильно";
			StringGridResults->Cells[2][i+1]="";
		}
		else{
			StringGridResults->Cells[1][i+1]="Ошибка";
		}
		StringGridResults->Cells[0][i+1]=IntToStr(i+1);
		StringGridResults->Cells[2][i+1]=TestQuestions[i].getReference();
	}

	ButtonGoBack->Top=StringGridResults->Top+StringGridResults->Height+3*INDENTION;
	ButtonCreateDocument->Top=ButtonGoBack->Top;
	if (isRandomTest==true) {
		ButtonGoBack->Left=(FTest->ClientWidth-(ButtonCreateDocument->Width+3*INDENTION+ButtonGoBack->Width))/2;
		ButtonCreateDocument->Left=ButtonGoBack->Left+ButtonGoBack->Width+3*INDENTION;
		ButtonSaveTest->Top=ButtonGoBack->Top+ButtonGoBack->Height+3*INDENTION;
		EditTestName->Top=ButtonGoBack->Top+ButtonGoBack->Height+3*INDENTION+ButtonSaveTest->Height-EditTestName->Height;
		EditTestName->Left=ButtonGoBack->Left;
		ButtonSaveTest->Left=EditTestName->Left+EditTestName->Width+3*INDENTION;
		//AnsiString name="Testing "+dateNow();
		//name = findFiles(name,"\\Testings",".xlsx");
		//EditTestName->Text=name;
		ButtonSaveTest->Caption="Сохранить тест как\n" + EditTestName->Text;
	}
	else{
		ButtonGoBack->Left=(FTest->ClientWidth - ButtonGoBack->Width - ButtonCreateDocument->Width - 3*INDENTION)/2;
		ButtonCreateDocument->Left=ButtonGoBack->Left+ButtonGoBack->Width+3*INDENTION;
	}

	if (wasAnyQuestionAnswered==true) saveResults();

	LabelEnd->Visible=true;
	LabelResult->Visible=true;
	LabelMark->Visible=true;
	StringGridResults->Visible=true;
	ButtonGoBack->Visible=true;
	ButtonCreateDocument->Visible=true;
	if (isRandomTest==true) {
		ButtonSaveTest->Visible=true;
		EditTestName->Visible=true;
	}
	if (wasSettingsChanged==true) {
		getSettings();
		wasSettingsChanged=false;
	}
	MemoQuestionText->Visible=false;
	globalCounter=-1;
	isAllowedToShow = true;
}

void __fastcall TFTest::ButtonAnswerClick(TObject *Sender)
{
	bool wasAnswered=false;
	AnsiString answer;
	int i;

	if ((TestQuestions[globalCounter].getQuestionType()=="choose" || TestQuestions[globalCounter].getQuestionType()=="music" || TestQuestions[globalCounter].getQuestionType()=="video") && RadioGroupVariants->ItemIndex != -1){
		FTest->MemoQuestionText->Visible=false;
		FTest->RadioGroupVariants->Visible=false;
		int choosedVariant=RadioGroupVariants->ItemIndex;
		answer=RadioGroupVariants->Items->Strings[choosedVariant];
		wasAnswered=true;
	}
	else if (TestQuestions[globalCounter].getQuestionType()=="write" && EditAnswer->Text!="") {
		EditAnswer->Visible=false;
		answer=EditAnswer->Text;
		wasAnswered=true;
	}
	else if (TestQuestions[globalCounter].getQuestionType()=="arrange") {
		wasAnswered=true;
		for (i = 0; i<ARRANGES; i++) {
			if (comboBoxAnswers[i]->Text=="Выберите ответ") {
				wasAnswered=false;
				comboBoxAnswers[i]->Color=clRed;
				TimerForWaiting->Interval=200;
				TimerForWaiting->Enabled=true;
				break;
			}
		}
		if (wasAnswered==false) {
			return;
		}

		bool isRight=true;
		if (FTest->ButtonUp->Visible==true) {
			for (i = 0; i < ARRANGES; i++) {
				if (panelsAnswers[i]->Caption != TestQuestions[globalCounter].getVariant(i+ARRANGES)) {
					isRight=false;
				}
				TestQuestions[globalCounter].setUserArrangeAnswer(i,panelsAnswers[i]->Caption);
			}
		}
		else if (comboBoxAnswers[0]->Visible==true) {
			for (i = 0; i < ARRANGES; i++) {
				if (comboBoxAnswers[i]->Text != TestQuestions[globalCounter].getVariant(i+ARRANGES)) {
					isRight=false;
				}
				TestQuestions[globalCounter].setUserArrangeAnswer(i,comboBoxAnswers[i]->Text);
			}
		}
		TestQuestions[globalCounter].setIsRightAnswer(isRight);
	}
	else if (TestQuestions[globalCounter].getQuestionType()=="somevariants") {
		wasAnswered=false;
		for (i = 0; i < variantsAmount+1; i++) {
			if (checkBoxes[i]->Checked==true) {
				wasAnswered=true;
				break;
			}
		}
		if (wasAnswered==false) {
			ButtonAnswer->Caption="Выберите ответы!";
			TimerForWaiting->Interval=1000;
			TimerForWaiting->Enabled=true;
		}
	}
	else if (TestQuestions[globalCounter].getQuestionType()=="arrangepictures") {
		wasAnswered=true;
		for (i = 0; i<ARRANGES; i++) {
			if (comboBoxAnswers[i]->Text=="Выберите ответ") {
				wasAnswered=false;
				comboBoxAnswers[i]->Color=clRed;
				TimerForWaiting->Interval=200;
				TimerForWaiting->Enabled=true;
				break;
			}
		}
		if (wasAnswered==false) {
			return;
		}

		bool isRight=true;
		if (FTest->ButtonRight->Visible==true) {
			for (i = 0; i < TestQuestions[globalCounter].getRightVariantsAmount(); i++) {
				if (panelsAnswers[i]->Caption != TestQuestions[globalCounter].getVariant(i)) {
					isRight=false;
				}
				TestQuestions[globalCounter].setUserArrangeAnswer(i,panelsAnswers[i]->Caption);
			}
		}
		else if (comboBoxAnswers[0]->Visible==true){
			for (i = 0; i < TestQuestions[globalCounter].getRightVariantsAmount(); i++) {
				if (comboBoxAnswers[i]->Text != TestQuestions[globalCounter].getVariant(i)) {
					isRight=false;
				}
				TestQuestions[globalCounter].setUserArrangeAnswer(i,comboBoxAnswers[i]->Text);
			}
        }
		TestQuestions[globalCounter].setIsRightAnswer(isRight);
		wasAnswered=true;
	}
	else {
		if (TestQuestions[globalCounter].getQuestionType()=="choose" || TestQuestions[globalCounter].getQuestionType()=="video" || TestQuestions[globalCounter].getQuestionType()=="music") ButtonAnswer->Caption="Выберите ответ!";
		else if (TestQuestions[globalCounter].getQuestionType()=="write") ButtonAnswer->Caption="Введите ответ!";
		wasAnswered=false;
	}

	if (wasAnswered==false) {
		TimerForWaiting->Interval=1000;
		TimerForWaiting->Enabled=true;
		return;
	}
	else wasAnyQuestionAnswered=true;

    if (FCharacteristics->CheckBoxTime->Checked==true) {
		TimerTime->Enabled=false;
		PanelTime->Font->Color=clBlack;
		seconds+=SECONDS_FOR_ANSWER-secondsForAnswer;
		PanelTime->Caption=timeToString(SECONDS_FOR_ANSWER);
		secondsForAnswer=SECONDS_FOR_ANSWER;
	}

	TestQuestions[globalCounter].setIsAnswered(wasAnswered);

	if (wasAnswered==true) {
		FTest->ProgressBar1->StepIt();
		testStatistics.addAmountOfAnsweredQuestions();

		TestQuestions[globalCounter].setUserAnswer(answer);
		if (TestQuestions[globalCounter].getQuestionType()=="choose" || TestQuestions[globalCounter].getQuestionType()=="write" || TestQuestions[globalCounter].getQuestionType()=="music" || TestQuestions[globalCounter].getQuestionType()=="video") {
			AnsiString rightAnswer=TestQuestions[globalCounter].getRightVariant();
			rightAnswer=rightAnswer.Trim();
			rightAnswer=rightAnswer.LowerCase();
			answer=answer.Trim();
			answer=answer.LowerCase();
			if (answer == rightAnswer) {
				TestQuestions[globalCounter].setIsRightAnswer(true);
				testStatistics.addCollectedPoints(TestQuestions[globalCounter].getDifficulty());
			}
			else {
				testStatistics.addWrongAnswers(IntToStr(globalCounter+1));
				TestQuestions[globalCounter].setIsRightAnswer(false);
			}
		}
		else if(TestQuestions[globalCounter].getQuestionType()=="arrange" || TestQuestions[globalCounter].getQuestionType()=="arrangepictures") {
			if (TestQuestions[globalCounter].getIsRightAnswer()==true) {
				testStatistics.addCollectedPoints(TestQuestions[globalCounter].getDifficulty());
			}
			else {
				testStatistics.addWrongAnswers(IntToStr(globalCounter+1));
			}
		}
		else if (TestQuestions[globalCounter].getQuestionType()=="somevariants") {
			int i,j,checkedAmount=0;
			bool isRight=true;
			AnsiString temp;
			for (i = 0; i < variantsAmount+1; i++) {
				if (checkBoxes[i]->Checked==true) {
					checkedAmount++;
					answer+=checkBoxes[i]->Caption+", ";
					bool tempIsRight=false;
					for (j=0; j < TestQuestions[globalCounter].getRightVariantsAmount(); j++) {
						if (TestQuestions[globalCounter].getVariant(j)==checkBoxes[i]->Caption) {
							tempIsRight=true;
							break;
						}
					}
					if (tempIsRight==false) {
						isRight=false;
					}
				}
			}
			if (checkedAmount!=TestQuestions[globalCounter].getRightVariantsAmount()) {
				isRight=false;
			}
			if (isRight==true) {
				testStatistics.addCollectedPoints(TestQuestions[globalCounter].getDifficulty());
			}
			else testStatistics.addWrongAnswers(IntToStr(globalCounter+1));
			TestQuestions[globalCounter].setIsRightAnswer(isRight);
			TestQuestions[globalCounter].setUserAnswer(answer);
		}
	}

	bool wasChanged=changeToNotAnsweredQuestion();

	if (wasChanged==false){
		MemoQuestionText->Visible=false;
		stopTest();
	}
}
//---------------------------------------------------------------------------
void changeQuestion(int index){
	int i,j, imagesAmount=0;

	FTest->MediaPlayer1->Enabled=true;
	FTest->ScrollBar1->Enabled=true;
	FTest->MemoQuestionText->Lines->Clear();
	FTest->RadioGroupVariants->Items->Clear();
	FTest->EditAnswer->Clear();
	FTest->MemoQuestionText->Visible=false;
	FTest->RadioGroupVariants->Visible=false;
	FTest->EditAnswer->Visible=false;
	FTest->LabelEnterAnswer->Visible=false;
	FTest->ButtonUp->Visible=false;
	FTest->ButtonDown->Visible=false;
	FTest->ButtonUpGeneral->Visible=false;
	FTest->ButtonDownGeneral->Visible=false;
	FTest->ButtonLeft->Visible=false;
	FTest->ButtonRight->Visible=false;
	FTest->ButtonLeftGeneral->Visible=false;
	FTest->ButtonRightGeneral->Visible=false;
	FTest->ScrollBar1->Visible=false;
	FTest->PanelVideo->Visible=false;
	FTest->MediaPlayer1->Visible=false;
	FTest->TimerForMusic->Enabled=false;
	FTest->ScrollBar1->Position=0;

	if (isMediaPlayerOpened==true) {
		FTest->MediaPlayer1->Close();
		isMediaPlayerOpened=false;
	}

	if (areImages==true || (isRandomTest==false && areImagesInCreated==true)) {
		for (i = 0; i < IMAGES_AMOUNT; i++) {
			if (isRandomTest==false && areImagesInCreated==false) {
				break;
			}
			images[i]->Visible=false;
		}
	}
	for (i = 0; i < ARRANGES; i++) {
		panelsInfo[i]->Visible=false;
		panelsAnswers[i]->Visible=false;
		comboBoxAnswers[i]->Visible=false;
	}
	for (i = 0; i < variantsAmount+1; i++) {
		checkBoxes[i]->Visible=false;
		if (checkBoxes[i]->Checked==true) {
			checkBoxes[i]->Checked=false;
		}
	}

	FTest->ButtonAnswer->Left=(FTest->ClientWidth - FTest->ButtonAnswer->Width)/2;
	FTest->MemoQuestionText->ShowHint=true;
	if (TestQuestions[index].getQuestionType()=="choose")
		FTest->MemoQuestionText->Hint="Нажмите на выбранный вариант.";
	else if (TestQuestions[index].getQuestionType()=="write")
		FTest->MemoQuestionText->Hint="Напишите ответ и нажмите \"Ответить\". \nСоблюдайте правила орфографии. Ответ требуется в именительном падеже. \nИспользуйте формы множественного и единственного числа.";
	else if (TestQuestions[index].getQuestionType()=="arrange"){
		if (FTest->ButtonUp->Visible==true || FTest->ButtonRight->Visible==true) {
			FTest->MemoQuestionText->Hint="Соотнесите элементы правой колонки с соответствующими элементами левой колонки. \nИспользуйте зеленые стрелки для перемены местами рядом стоящих элементов. \nИспользуйте синие стрелки для передвижения по элементам.";
		}
		else FTest->MemoQuestionText->Hint="Выберите ответ из списка";
	}
	else if (TestQuestions[index].getQuestionType()=="somevariants")
		FTest->MemoQuestionText->Hint="Выберите верные варианты.";
	else if (TestQuestions[index].getQuestionType()=="arrangepictures"){
		if (FTest->ButtonUp->Visible==true || FTest->ButtonRight->Visible==true){
			FTest->MemoQuestionText->Hint="Соотнесите изображения и ответы. \nИспользуйте зеленые стрелки для перемены местами рядом стоящих ответов. \nИспользуйте синие стрелки для передвижения по ответам.";
		}
		else FTest->MemoQuestionText->Hint="Выберите ответ из списка";
	}
											   //ShowMessage(TestQuestions[globalCounter].getText());
	FTest->MemoQuestionText->Lines->Add(TestQuestions[globalCounter].getText());

	if (TestQuestions[index].getQuestionType()=="choose" || TestQuestions[index].getQuestionType()=="write" || TestQuestions[index].getQuestionType()=="music" || TestQuestions[index].getQuestionType()=="video"){
		for (i = 0; i < variantsAmount; i++) {
			FTest->RadioGroupVariants->Items->Add(TestQuestions[index].getVariant(i));
		}
		FTest->RadioGroupVariants->Items->Insert(random(variantsAmount),TestQuestions[index].getRightVariant());
	}
				   //ShowMessage(IntToStr(HEIGHT));
	FTest->MemoQuestionText->Height=(FTest->MemoQuestionText->Lines->Count)*HEIGHT;

	if (TestQuestions[index].getPicture(0)!="-" && (areImages==true || (isRandomTest==false && areImagesInCreated==true))){
		if (isRandomTest==false && areImagesInCreated==false) {
			goto aaa;
		}
		AnsiString temp;
		for (i = 0; i < IMAGES_AMOUNT; i++) {
			if (TestQuestions[index].getPicture(i)!="-") {
				imagesAmount++;
			}
			else break;
		}

		TestQuestions[index].setImagesAmount(imagesAmount);

		for (i = 0; i < imagesAmount; i++) {
			images[i]->Stretch=true;
			images[i]->Top=FTest->MemoQuestionText->Height+FTest->MemoQuestionText->Top+3*INDENTION;
			images[i]->Height=IMAGE_HEIGHT;
			images[i]->Width=IMAGE_WIDTH;
			try{
				images[i]->Picture->LoadFromFile(TestQuestions[index].getPicture(i));
			}
			catch(...){
				FTest->MemoQuestionText->Lines->Add("");
				FTest->MemoQuestionText->Lines->Add("Возникла ошибка при попытке загрузки файла "+TestQuestions[index].getPicture(i));
            	FTest->MemoQuestionText->Perform(EM_SCROLL,SB_LINEDOWN,0);
			}
			images[i]->Visible=true;
			images[i]->Parent=FTest;
			images[i]->Hint="Нажмите, чтобы увеличить изображение";
			images[i]->ShowHint=true;
			images[i]->Cursor=crHandPoint;
			images[i]->OnClick=FTest->Image1->OnClick;
		}
		FTest->RadioGroupVariants->Top=images[0]->Top+IMAGE_HEIGHT+3*INDENTION;
	}
	else {
		aaa:
		FTest->RadioGroupVariants->Top=FTest->MemoQuestionText->Top+FTest->MemoQuestionText->Height+3*INDENTION;
		TestQuestions[globalCounter].setImagesAmount(imagesAmount);
	}

	int distanceBetweenPictures=(FTest->MemoQuestionText->Width-IMAGE_WIDTH*imagesAmount)/(imagesAmount+1);
	for (i = 0; i < imagesAmount; i++) {
		if (imagesAmount!=4) {
			images[i]->Left=distanceBetweenPictures+FTest->MemoQuestionText->Left + i*(IMAGE_WIDTH+distanceBetweenPictures);
		}
		else images[i]->Left= FTest->MemoQuestionText->Left + i*(IMAGE_WIDTH+distanceBetweenPictures);
	}
	if (imagesAmount==1) {
		images[0]->Left=(FTest->ClientWidth - images[0]->Width)/2;
	}

	FTest->RadioGroupVariants->Left=FTest->MemoQuestionText->Left;
	FTest->MemoQuestionText->Visible=true;
	if (TestQuestions[index].getQuestionType()=="choose") {
		FTest->RadioGroupVariants->Columns=2;
		if (variantsAmount%2==0) { //в начале variantsAmount-1!
			FTest->RadioGroupVariants->Height=INDENTION*5*(((variantsAmount+2)/2)+1);
		}
		else FTest->RadioGroupVariants->Height=INDENTION*5*(((variantsAmount+1)/2)+1);

		FTest->ButtonAnswer->Top=FTest->RadioGroupVariants->Height+FTest->RadioGroupVariants->Top+3*INDENTION;
		FTest->RadioGroupVariants->Visible=true;
	}
	else if(TestQuestions[index].getQuestionType()=="write"){
		FTest->LabelEnterAnswer->Top=FTest->RadioGroupVariants->Top;
		FTest->EditAnswer->Top=FTest->LabelEnterAnswer->Top;
		FTest->EditAnswer->Left=(FTest->ClientWidth - FTest->EditAnswer->Width)/2;
		FTest->LabelEnterAnswer->Left=FTest->EditAnswer->Left - FTest->LabelEnterAnswer->Width;
		FTest->ButtonAnswer->Top=FTest->EditAnswer->Height+FTest->EditAnswer->Top+3*INDENTION;

		FTest->LabelEnterAnswer->Visible=true;
		FTest->EditAnswer->Visible=true;
		FTest->EditAnswer->SetFocus();
		UnicodeString Russian="00000419";
		LoadKeyboardLayout(Russian.w_str(), KLF_ACTIVATE);
	}
	else if (TestQuestions[index].getQuestionType()=="arrange") {
		int* usedPositions;
		int position;
		usedPositions=new int[ARRANGES];
		zeroingArray(usedPositions,ARRANGES,-1);
		for (i = 0; i < ARRANGES; i++) {
			panelsInfo[i]->Caption=TestQuestions[globalCounter].getVariant(i);
			panelsInfo[i]->Left=FTest->RadioGroupVariants->Left+10*INDENTION;
			panelsInfo[i]->Width=WIDTH_INFO;
			panelsInfo[i]->Height=HEIGHT;
			panelsInfo[i]->Top=FTest->RadioGroupVariants->Top+i*(HEIGHT+2*INDENTION);
			panelsInfo[i]->Font->Height=FONT_HEIGHT;
		}
		if (FTest->MenuArrow->Checked==true){
			for (i=ARRANGES; i < ARRANGES*2; i++) {
				panelsAnswers[i-ARRANGES]->Width=WIDTH_ANSWERS;
				panelsAnswers[i-ARRANGES]->Height=HEIGHT;
				panelsAnswers[i-ARRANGES]->Left=panelsInfo[i-ARRANGES]->Left+WIDTH_INFO+5*INDENTION;
				panelsAnswers[i-ARRANGES]->Top=FTest->RadioGroupVariants->Top+(i-ARRANGES)*(HEIGHT+2*INDENTION);
				panelsAnswers[i-ARRANGES]->Font->Height=FONT_HEIGHT;
				panelsAnswers[i-ARRANGES]->TabStop=true;
			}
			for (i = 0; i < ARRANGES; i++) {
				position=1;
				if (i==0) {
					while (position==1){
						position=generateRandomNumber(usedPositions,ARRANGES,ARRANGES);
					}
				}
				else position=generateRandomNumber(usedPositions,ARRANGES,ARRANGES);
				usedPositions[i]=position;
				panelsAnswers[position-1]->Caption=TestQuestions[index].getVariant(i+ARRANGES);
			}
			for (i = 0; i < ARRANGES; i++) {
				panelsAnswers[i]->Visible=true;
				panelsAnswers[i]->Parent=FTest;
			}
			upDown();
			FTest->ButtonUp->Visible=true;
			FTest->ButtonDown->Visible=true;
			FTest->ButtonUpGeneral->Visible=true;
			FTest->ButtonDownGeneral->Visible=true;
		}
		else if (FTest->MenuDropDown->Checked==true){
			for (i=ARRANGES; i < ARRANGES*2; i++) {
				comboBoxAnswers[i-ARRANGES]->Width=WIDTH_ANSWERS;
				comboBoxAnswers[i-ARRANGES]->Height=HEIGHT;
				comboBoxAnswers[i-ARRANGES]->Left=panelsInfo[i-ARRANGES]->Left+WIDTH_INFO+5*INDENTION;
				comboBoxAnswers[i-ARRANGES]->Top=FTest->RadioGroupVariants->Top+(i-ARRANGES)*(HEIGHT+2*INDENTION);
				comboBoxAnswers[i-ARRANGES]->Font->Height=FONT_HEIGHT;
				comboBoxAnswers[i-ARRANGES]->TabStop=true;
				comboBoxAnswers[i-ARRANGES]->Color=FTest->Color;
			}
			for (i = 0; i < ARRANGES; i++) {
				comboBoxAnswers[i]->Parent=FTest;
				comboBoxAnswers[i]->Clear();
				for (j=0; j < ARRANGES; j++) {
					if (i==0) {
						position=generateRandomNumber(usedPositions,ARRANGES,ARRANGES);
						usedPositions[j]=position;
						comboBoxAnswers[i]->Items->Add(TestQuestions[index].getVariant(position-1+ARRANGES));
					}
					else {
						comboBoxAnswers[i]->Items->Add(TestQuestions[index].getVariant(usedPositions[j]-1+ARRANGES));
					}
				}
			}
			for (i = 0; i < ARRANGES; i++) {
				comboBoxAnswers[i]->Visible=true;
				comboBoxAnswers[i]->Text="Выберите ответ";
			}
		}
		for (i = 0; i < ARRANGES; i++) {
			panelsInfo[i]->Visible=true;
			panelsInfo[i]->Parent=FTest;
		}
		FTest->ButtonAnswer->Top=panelsInfo[ARRANGES-1]->Top+panelsInfo[ARRANGES-1]->Height+3*INDENTION;
		delete []usedPositions;
	}
	else if (TestQuestions[index].getQuestionType()=="somevariants") {
		int* usedPositions;
		int position;
		usedPositions=new int[variantsAmount+1];
		zeroingArray(usedPositions,variantsAmount+1,-1);
		for (i = 0; i < variantsAmount+1 ; i++) {
			position=generateRandomNumber(usedPositions,variantsAmount+1,variantsAmount+1);
			usedPositions[i]=position;
			checkBoxes[position-1]->Caption=TestQuestions[index].getVariant(i);
		}
		for (i = 0; i < variantsAmount+1; i++) {
			checkBoxes[i]->Font->Height=FONT_HEIGHT;
			if (i%3==0)
				 checkBoxes[i]->Left=FTest->RadioGroupVariants->Left+10*INDENTION;
			else if (i%3==1)
				 checkBoxes[i]->Left=FTest->RadioGroupVariants->Left+FTest->RadioGroupVariants->Width/3+10*INDENTION;
			else checkBoxes[i]->Left=FTest->RadioGroupVariants->Left+2*FTest->RadioGroupVariants->Width/3+10*INDENTION;
			checkBoxes[i]->Top=FTest->RadioGroupVariants->Top+(i/3)*(HEIGHT+2*INDENTION);
			checkBoxes[i]->Height=HEIGHT;
			checkBoxes[i]->Width=WIDTH_ANSWERS;
			checkBoxes[i]->Visible=true;
			checkBoxes[i]->Parent=FTest;
		}
		FTest->ButtonAnswer->Top=checkBoxes[variantsAmount]->Top+checkBoxes[variantsAmount]->Height+3*INDENTION;
		delete []usedPositions;
	}
	else if (TestQuestions[index].getQuestionType()=="arrangepictures") {
		int* usedPositions;
		int position;
		usedPositions=new int[TestQuestions[index].getRightVariantsAmount()];
		zeroingArray(usedPositions,TestQuestions[index].getRightVariantsAmount(),-1);

		if (FTest->MenuArrow->Checked==true) {
			for (i=0; i < TestQuestions[index].getRightVariantsAmount(); i++) {
				panelsAnswers[i]->Width=WIDTH_ANSWERS;
				panelsAnswers[i]->Height=HEIGHT;
				panelsAnswers[i]->Left=images[i]->Left;
				panelsAnswers[i]->Width=images[i]->Width;
				panelsAnswers[i]->Top=images[i]->Top+images[i]->Height+2*INDENTION;
				panelsAnswers[i]->Font->Height=FONT_HEIGHT;
				panelsAnswers[i]->TabStop=true;
			}

			for (i = 0; i < TestQuestions[globalCounter].getRightVariantsAmount(); i++) {
				position=1;
				if (i==0) {
					while (position==1){
						position=generateRandomNumber(usedPositions,TestQuestions[index].getRightVariantsAmount(),TestQuestions[index].getRightVariantsAmount());
					}
				}
				else position=generateRandomNumber(usedPositions,TestQuestions[globalCounter].getRightVariantsAmount(),TestQuestions[index].getRightVariantsAmount());
				usedPositions[i]=position;
				panelsAnswers[position-1]->Caption=TestQuestions[index].getVariant(i);
			}
			for (i = 0; i < TestQuestions[index].getRightVariantsAmount(); i++) {
				panelsAnswers[i]->Visible=true;
				panelsAnswers[i]->Parent=FTest;
			}
			leftRight();
			FTest->ButtonLeft->Visible=true;
			FTest->ButtonRight->Visible=true;
			FTest->ButtonLeftGeneral->Visible=true;
			FTest->ButtonRightGeneral->Visible=true;
			FTest->ButtonAnswer->Top=FTest->ButtonLeftGeneral->Top+FTest->ButtonLeftGeneral->Height+5*INDENTION;
		}

		else if (FTest->MenuDropDown->Checked==true){
			for (i=0; i < TestQuestions[index].getRightVariantsAmount(); i++) {
				comboBoxAnswers[i]->Width=WIDTH_ANSWERS;
				comboBoxAnswers[i]->Height=HEIGHT;
				comboBoxAnswers[i]->Left=images[i]->Left;
				comboBoxAnswers[i]->Width=images[i]->Width;
				comboBoxAnswers[i]->Top=images[i]->Top+images[i]->Height+2*INDENTION;
				comboBoxAnswers[i]->Font->Height=FONT_HEIGHT;
				comboBoxAnswers[i]->TabStop=true;
				comboBoxAnswers[i]->Color=FTest->Color;
			}

			for (i = 0; i < TestQuestions[index].getRightVariantsAmount(); i++) {
				comboBoxAnswers[i]->Parent=FTest;
				comboBoxAnswers[i]->Clear();
				for (j=0; j < TestQuestions[index].getRightVariantsAmount(); j++) {
					if (i==0) {
						position=generateRandomNumber(usedPositions,TestQuestions[index].getRightVariantsAmount(),TestQuestions[index].getRightVariantsAmount());
						usedPositions[j]=position;
						comboBoxAnswers[i]->Items->Add(TestQuestions[index].getVariant(position-1));
					}
					else {
						comboBoxAnswers[i]->Items->Add(TestQuestions[index].getVariant(usedPositions[j]-1));
					}
				}
			}
			for (i = 0; i < TestQuestions[index].getRightVariantsAmount(); i++) {
				comboBoxAnswers[i]->Visible=true;
				comboBoxAnswers[i]->Text="Выберите ответ";
			}
			FTest->ButtonAnswer->Top=comboBoxAnswers[0]->Top+comboBoxAnswers[0]->Height+5*INDENTION;
		}

		delete []usedPositions;
	}
	else if (TestQuestions[index].getQuestionType()=="music"){
		try{
			int imagesAmount=TestQuestions[index].getImagesAmount();
			FTest->MediaPlayer1->Width=19*INDENTION;
			FTest->MediaPlayer1->Top=FTest->MemoQuestionText->Top+FTest->MemoQuestionText->Height+4*INDENTION;
			FTest->ScrollBar1->Width=IMAGE_WIDTH;
			if (imagesAmount==0) {
				FTest->MediaPlayer1->Left=(FTest->ClientWidth-FTest->MediaPlayer1->Width)/2;
				FTest->ScrollBar1->Left=(FTest->ClientWidth - FTest->ScrollBar1->Width)/2;
				FTest->ScrollBar1->Top=FTest->MediaPlayer1->Top+FTest->MediaPlayer1->Height+4*INDENTION;
				FTest->RadioGroupVariants->Top=FTest->ScrollBar1->Top+FTest->ScrollBar1->Height+4*INDENTION;
			}
			else {
				for (i = 0; i < imagesAmount; i++) {
					images[i]->Left-=images[i]->Width/2;
				}
				FTest->ScrollBar1->Left=images[imagesAmount-1]->Left + images[imagesAmount-1]->Width + 4*INDENTION;
				FTest->MediaPlayer1->Left=FTest->ScrollBar1->Left + (FTest->ScrollBar1->Width - FTest->MediaPlayer1->Width)/2;
				FTest->MediaPlayer1->Top=images[imagesAmount-1]->Top + (images[imagesAmount-1]->Height-(FTest->MediaPlayer1->Height+FTest->ScrollBar1->Height+4*INDENTION))/2;
				FTest->ScrollBar1->Top=FTest->MediaPlayer1->Top+FTest->MediaPlayer1->Height+4*INDENTION;
				FTest->RadioGroupVariants->Top=images[imagesAmount-1]->Top + images[imagesAmount-1]->Height+4*INDENTION;
			}

			FTest->RadioGroupVariants->Columns=2;
			if (variantsAmount%2==0) { //в начале variantsAmount-1!
				FTest->RadioGroupVariants->Height=INDENTION*5*(((variantsAmount+2)/2)+1);
			}
			else FTest->RadioGroupVariants->Height=INDENTION*5*(((variantsAmount+1)/2)+1);

			FTest->ButtonAnswer->Top=FTest->RadioGroupVariants->Top+FTest->RadioGroupVariants->Height+4*INDENTION;

			FTest->MediaPlayer1->Visible=true;
			FTest->RadioGroupVariants->Visible=true;
			FTest->ScrollBar1->Visible=true;
			try{
				FTest->MediaPlayer1->FileName=TestQuestions[index].getWayToMediaFile();
				FTest->MediaPlayer1->Open();
				isMediaPlayerOpened=true;
			}
			catch(...){
				Application->Title="Ошибка";
				FTest->MediaPlayer1->Enabled=false;
				FTest->ScrollBar1->Enabled=false;
				FTest->ButtonAnswer->Top=FTest->RadioGroupVariants->Top+FTest->RadioGroupVariants->Height+4*INDENTION;
				Application->Title="Ошибка";
				ShowMessage("Ошибка при открытии файла "+TestQuestions[index].getWayToMediaFile());
			}

			FTest->ScrollBar1->Max=FTest->MediaPlayer1->Length/1000;
		}
		catch(...){
			FTest->MediaPlayer1->Enabled=false;
			FTest->ScrollBar1->Enabled=false;
			FTest->ButtonAnswer->Top=FTest->RadioGroupVariants->Top+FTest->RadioGroupVariants->Height+4*INDENTION;
		}
	}
	else if (TestQuestions[index].getQuestionType()=="video"){
		try{
			FTest->PanelVideo->Width=TestQuestions[index].getVideoResolutionX();
			FTest->PanelVideo->Height=TestQuestions[index].getVideoResolutionY();
			FTest->PanelVideo->Left=FTest->MemoQuestionText->Left;
			FTest->PanelVideo->Top=FTest->MemoQuestionText->Top+FTest->MemoQuestionText->Height+4*INDENTION;

			if (FTest->PanelVideo->Width > 480) {
				FTest->RadioGroupVariants->Columns=1;
				FTest->RadioGroupVariants->Height=INDENTION*5*(variantsAmount+1);
			}
			else{
				FTest->RadioGroupVariants->Columns=2;
				if (variantsAmount%2==0) { //в начале variantsAmount-1!
					FTest->RadioGroupVariants->Height=INDENTION*5*(((variantsAmount+2)/2)+1);
				}
				else FTest->RadioGroupVariants->Height=INDENTION*5*(((variantsAmount+1)/2)+1);
			}

			FTest->RadioGroupVariants->Left=FTest->PanelVideo->Left + FTest->PanelVideo->Width + 4*INDENTION;
			FTest->RadioGroupVariants->Top=FTest->PanelVideo->Top;

			FTest->MediaPlayer1->Width=19*INDENTION;
			FTest->MediaPlayer1->Left=FTest->MemoQuestionText->Left + (FTest->PanelVideo->Width - FTest->MediaPlayer1->Width)/2;
			FTest->MediaPlayer1->Top=FTest->PanelVideo->Top + FTest->PanelVideo->Height + 4*INDENTION;

			FTest->MediaPlayer1->Visible=true;
			FTest->PanelVideo->Visible=true;
			FTest->RadioGroupVariants->Visible=true;
			FTest->ScrollBar1->Visible=true;

			try{
				FTest->MediaPlayer1->FileName=TestQuestions[index].getWayToMediaFile();
				FTest->MediaPlayer1->Open();
				isMediaPlayerOpened=true;
				FTest->MediaPlayer1->Display=FTest->PanelVideo;
			}
			catch(...){
				FTest->MediaPlayer1->Enabled=false;
				FTest->ScrollBar1->Enabled=false;
				FTest->ButtonAnswer->Top=FTest->RadioGroupVariants->Top+FTest->RadioGroupVariants->Height+4*INDENTION;
				Application->Title="Ошибка";
				ShowMessage("Ошибка при открытии файла "+TestQuestions[index].getWayToMediaFile());
			}

			FTest->ScrollBar1->Width=FTest->PanelVideo->Width;
			FTest->ScrollBar1->Left=FTest->MemoQuestionText->Left;
			FTest->ScrollBar1->Top=FTest->MediaPlayer1->Top + FTest->MediaPlayer1->Height + 3*INDENTION;

			FTest->ScrollBar1->Max=FTest->MediaPlayer1->Length/1000;

			FTest->ButtonAnswer->Top=FTest->RadioGroupVariants->Top+FTest->RadioGroupVariants->Height+4*INDENTION;
			if (FTest->PanelVideo->Left+FTest->PanelVideo->Width >= FTest->ButtonAnswer->Left){
				FTest->ButtonAnswer->Left=FTest->PanelVideo->Left+FTest->PanelVideo->Width+3*INDENTION;
			}
			if (FTest->ScrollBar1->Top+FTest->ScrollBar1->Height > FTest->ButtonPreviousQuestion->Top) {
				FTest->MediaPlayer1->Top=FTest->PanelVideo->Top+FTest->PanelVideo->Height+INDENTION;
				FTest->ScrollBar1->Top=FTest->MediaPlayer1->Top+FTest->MediaPlayer1->Height+INDENTION;
			}
		}
		catch(...){
			FTest->MediaPlayer1->Enabled=false;
			FTest->ScrollBar1->Enabled=false;
			FTest->ButtonAnswer->Top=FTest->RadioGroupVariants->Top+FTest->RadioGroupVariants->Height+4*INDENTION;
		}
	}

	FTest->LabelBottomLine->Left=FTest->ButtonPreviousQuestion->Left + FTest->ButtonPreviousQuestion->Width + (FTest->ClientWidth - FTest->ButtonPreviousQuestion->Width - FTest->ButtonNextQuestion->Width - FTest->LabelBottomLine->Width - 4*INDENTION)/2;
	FTest->LabelBottomLine->Top=FTest->ButtonNextQuestion->Top + (FTest->ButtonNextQuestion->Height - FTest->LabelBottomLine->Height)/2;
	FTest->LabelBottomLine->Caption="Вопрос "+IntToStr(globalCounter+1)+" из "+IntToStr(questionsAmount);
    if ((FTest->LabelBottomLine->Top - FTest->ButtonAnswer->Top) <= FTest->ButtonAnswer->Height) {
		FTest->LabelBottomLine->Left=FTest->ButtonPreviousQuestion->Left + FTest->ButtonPreviousQuestion->Width + 2*INDENTION;
		FTest->ButtonAnswer->Left=FTest->LabelBottomLine->Left + FTest->LabelBottomLine->Width+2*INDENTION;
	}
	FTest->RadioGroupVariants->Width=FTest->MemoQuestionText->Width-FTest->RadioGroupVariants->Left;
	if (FTest->ButtonAnswer->Top+FTest->ButtonAnswer->Height > FTest->ClientHeight) {
		FTest->ButtonAnswer->Top=FTest->ClientHeight-FTest->ButtonAnswer->Height;
	}

	if (FCharacteristics->CheckBoxTime->Checked==true) {
		FTest->TimerTime->Enabled=true;
	}
}

void deleteTest(){
	int i;
	for (i = 0; i < IMAGES_AMOUNT; i++) {
		images[i]->Free();
	}
	for (i = ARRANGES-1; i >=0; i--) {
		delete panelsInfo[i];
		delete panelsAnswers[i];
		delete comboBoxAnswers[i];
	}
	for (i = MAX_VARIANTS_AMOUNT-1; i >=0; i--) {
		delete checkBoxes[i];
	}
	testStatistics.clearing();
	delete []TestQuestions;
}

void __fastcall TFTest::ButtonGoBackClick(TObject *Sender)
{
	deleteTest();
	FTest->Hide();
	FStart->Show();
}
//---------------------------------------------------------------------------

void addParagraph(int* paragraphCounter,Variant* Paragraphs,AnsiString text){
	Variant Paragraph;
	(*Paragraphs).OleProcedure("Add");
	Paragraph=(*Paragraphs).OleFunction("Item",*paragraphCounter);
	Paragraph.OlePropertyGet("Range").OlePropertySet("Text",StringToOleStr(text));
	(*paragraphCounter)++;
}

AnsiString dateNow(){
	TDateTime Date;
	AnsiString date=FormatDateTime("d m y",Date.CurrentDateTime());
	int length=date.Length();
	for (int i = 1; i <= length; i++) {
		if (date[i]==' ') {
			date[i]='.';
		}
	}
	return date;
}

int creatingDocument(){
	FTest->MemoCheck->Lines->Clear();
	FTest->MemoCheck->Width=IMAGE_WIDTH+50;
	FTest->MemoCheck->Top=FTest->ProgressBar1->Top+4*INDENTION;
	FTest->MemoCheck->Left=(FTest->ClientWidth-FTest->MemoCheck->Width)/2;
	//FTest->MemoCheck->Visible=true;

	Variant WordApplication,WordDocuments,Document,Paragraphs,Paragraph,Image;
	TDateTime Date;
	int images;
	int i,j,paragraphsCounter=1;
	AnsiString title;
	AnsiString name;
	AnsiString fileName;

	try{
		FTest->MemoCheck->Lines->Add("Получение требуемого имени файла");
		title="Тестирование "+dateNow()+" \t\t\t\tТестируемый: " + username;
		name=username+" "+dateNow();
		fileName=CURRENT_DIRECTORY+"\\Documents\\"+findFiles(name,"\\Documents",".docx")+".docx";

		if (FCharacteristics->CheckBoxAutoCreateDocument->Checked == true) FTest->LabelQuestionNumber->Visible=true;
		FTest->LabelQuestionNumber->Caption="Подготовка";
		FTest->LabelQuestionNumber->Left=(FTest->ClientWidth-FTest->LabelQuestionNumber->Width)/2;
		FTest->MemoCheck->Lines->Add("Требуемое имя файла найдено");
	}
	catch(...){
		Application->Title="Ошибка";
		ShowMessage("Ошибка при инициализации данных");
		return 1;
	}

	try{
		FTest->MemoCheck->Lines->Add("Открытие Word");
		FTest->LabelQuestionNumber->Caption="Открытие Microsoft Word";
		FTest->LabelQuestionNumber->Left=(FTest->ClientWidth-FTest->LabelQuestionNumber->Width)/2;

		WordApplication=CreateOleObject("Word.Application");
		FTest->MemoCheck->Lines->Add("Word открыт");
	}
	catch(...){
		Application->Title="Ошибка";
		ShowMessage("Ошибка открытия Word");
		return 1;
	}

	try{
		FTest->MemoCheck->Lines->Add("Настройка документа Word");
		FTest->LabelQuestionNumber->Caption="Подготовка документа";
		FTest->LabelQuestionNumber->Left=(FTest->ClientWidth-FTest->LabelQuestionNumber->Width)/2;
		WordDocuments=WordApplication.OlePropertyGet("Documents");
		WordDocuments.OleProcedure("Add");
		Document=WordDocuments.OleFunction("Item",paragraphsCounter);
		paragraphsCounter++;
		Document.OleProcedure("Activate");
		Paragraphs=Document.OlePropertyGet("Paragraphs");
		WordApplication.OlePropertyGet("Options").OlePropertySet("CheckGrammarAsYouType",false);
		WordApplication.OlePropertyGet("Options").OlePropertySet("CheckGrammarWithSpelling",false);

		addParagraph(&paragraphsCounter,&Paragraphs,title);

		FTest->ProgressBar1->Width=FTest->ClientWidth-4*INDENTION;
		FTest->ProgressBar1->Max=questionsAmount;
		FTest->ProgressBar1->Step=1;
		FTest->ProgressBar1->Position=0;
		FTest->ProgressBar1->Visible=true;

		FTest->MemoCheck->Lines->Add("Word документ настроен");
		FTest->MemoCheck->Lines->Add("");
		FTest->MemoCheck->Lines->Add("Начало записи вопросов");

		for (i = 0; i <questionsAmount ; i++) {
			if (i>=150) {
				ShowMessage("Ошибка в цикле");
				break;
			}
			FTest->MemoCheck->Lines->Add(IntToStr(i+1)+" - "+"");

			try{
				FTest->ProgressBar1->StepIt();
				FTest->LabelQuestionNumber->Caption="Вопрос №"+IntToStr(i+1);
				FTest->LabelQuestionNumber->Left=(FTest->ClientWidth-FTest->LabelQuestionNumber->Width)/2;
				images=0;

					FTest->MemoCheck->Lines->Add(IntToStr(i+1)+" - "+"Заголовок записан");
				addParagraph(&paragraphsCounter,&Paragraphs,"");
				addParagraph(&paragraphsCounter,&Paragraphs,"Вопрос "+IntToStr(i+1)+": "+TestQuestions[i].getText());
				FTest->MemoCheck->Lines->Add(IntToStr(i+1)+" - "+"Текст вопроса записан");
			}
			catch(...){
				ShowMessage("Ошибка при записи формулировки вопроса");
				WordApplication.OleProcedure("Quit");
				return 1;
			}

			if (areImages==true || (isRandomTest==false && areImagesInCreated==true)) {
				if (isRandomTest==false && areImagesInCreated==false) {
					goto bbb;
				}
				try {
					FTest->MemoCheck->Lines->Add(IntToStr(i+1)+" - "+"Запись изображений");
					for (j=TestQuestions[i].getImagesAmount()-1; j>=0; j--) {
						if (j>=150 || j<0) {
							throw 0;
						}
								  FTest->MemoCheck->Lines->Add(IntToStr(i+1)+" - "+IntToStr(j)+"- изображение");
						if (TestQuestions[i].getPicture(j)!="-"){
							images++;
							if (images<2) {
								Paragraphs.OleProcedure("Add");
								Paragraph=Paragraphs.OleFunction("Item",paragraphsCounter);
							}
							Paragraph.OlePropertyGet("Range").OlePropertyGet("InlineShapes").OleProcedure("AddPicture", StringToOleStr(TestQuestions[i].getPicture(j)), false, true);
							int count=Paragraph.OlePropertyGet("Range").OlePropertyGet("InlineShapes").OlePropertyGet("Count");
							Image=Paragraph.OlePropertyGet("Range").OlePropertyGet("InlineShapes").OleFunction("Item",1);
							Image.OlePropertySet("Width",150);
							Image.OlePropertySet("Height",100);
						}
					}
					if (images!=0) paragraphsCounter++;
				}
				catch(...){
					ShowMessage("Ошибка при загрузке изображений в отчет по тестированию");
					WordApplication.OleProcedure("Quit");
					return 1;
				}
				bbb:
			}

			FTest->MemoCheck->Lines->Add(IntToStr(i+1)+" - "+"Изображения записаны");
			if (TestQuestions[i].getQuestionType()=="choose" || TestQuestions[i].getQuestionType()=="music" || TestQuestions[i].getQuestionType()=="video"){
				try{
					FTest->MemoCheck->Lines->Add(IntToStr(i+1)+" - "+"Тип вопроса - мультимедийный или с 1 правильным вариантом");

					int position;
					bool isAnswerWritten=false;
					position=1+generateRandomNumber(variantsAmount-1);
					FTest->MemoCheck->Lines->Add(IntToStr(i+1)+" - "+"Запись вариантов ответа,"+ " правильный на "+IntToStr(position)+" позиции");
					for (j = 0; j <variantsAmount+1; j++) {
						if (j==position-1) {
							addParagraph(&paragraphsCounter,&Paragraphs,"\t"+IntToStr(j+1)+". "+TestQuestions[i].getRightVariant());
							isAnswerWritten=true;
						}
						else {
							if (isAnswerWritten==false) {
								addParagraph(&paragraphsCounter,&Paragraphs,"\t"+IntToStr(j+1)+". "+TestQuestions[i].getVariant(j));
							}
							else addParagraph(&paragraphsCounter,&Paragraphs,"\t"+IntToStr(j+1)+". "+TestQuestions[i].getVariant(j-1));
						}
					}
				}
				catch(...){
					Application->Title="Ошибка";
					ShowMessage("Ошибка при записи вопроса "+TestQuestions[i].getText());
					WordApplication.OleProcedure("Quit");
					return 1;
				}
				FTest->MemoCheck->Lines->Add(IntToStr(i+1)+" - "+"Варианты записаны ");
			}
			else if (TestQuestions[i].getQuestionType()=="arrange") {
				FTest->MemoCheck->Lines->Add(IntToStr(i+1)+" - "+"Тип вопроса - \"расставить ответы\"");
				try{

					int *usedPositions=new int[ARRANGES];
					int position;
					zeroingArray(usedPositions,ARRANGES,-1);
					FTest->MemoCheck->Lines->Add(IntToStr(i+1)+" - "+"Запись вариантов");
					for (j = 0; j <ARRANGES; j++) {
						position=generateRandomNumber(usedPositions,ARRANGES,ARRANGES);
						FTest->MemoCheck->Lines->Add(IntToStr(i+1)+" - "+"Позиция: "+IntToStr(position));
						usedPositions[j]=position;
						addParagraph(&paragraphsCounter,&Paragraphs,"\t"+TestQuestions[i].getVariant(j)+"\t\t "+TestQuestions[i].getVariant(position-1+ARRANGES));
					}
					delete []usedPositions;
					FTest->MemoCheck->Lines->Add(IntToStr(i+1)+" - "+"варианты записаны");
				}
				catch(...){
					Application->Title="Ошибка";
					ShowMessage("Ошибка при записи вопроса "+TestQuestions[i].getText());
					WordApplication.OleProcedure("Quit");
					return 1;
				}
			}
			else if (TestQuestions[i].getQuestionType()=="arrangepictures") {
				FTest->MemoCheck->Lines->Add(IntToStr(i+1)+" - "+"Тип вопроса - \"Расставить ответы по изображениям\"");
				try {
					AnsiString variants="";
					int *usedPositions=new int[TestQuestions[i].getRightVariantsAmount()];
					FTest->MemoCheck->Lines->Add(IntToStr(i+1)+" - "+"Правильных вариантов: "+IntToStr(TestQuestions[i].getRightVariantsAmount()));
					int position;
					zeroingArray(usedPositions,TestQuestions[i].getRightVariantsAmount(),-1);
					for (j = 0; j <TestQuestions[i].getRightVariantsAmount(); j++) {
						position=generateRandomNumber(usedPositions,TestQuestions[i].getRightVariantsAmount(),TestQuestions[i].getRightVariantsAmount());
						variants+= TestQuestions[i].getVariant(position-1);
						usedPositions[j]=position;
							FTest->MemoCheck->Lines->Add(IntToStr(i+1)+" - "+"Позиция: "+IntToStr(position));
						if (TestQuestions[i].getRightVariantsAmount()-j>=2) {
							variants+=", ";
						}
						else variants+=".";
					}
					addParagraph(&paragraphsCounter,&Paragraphs,variants);
					delete []usedPositions;
						FTest->MemoCheck->Lines->Add(IntToStr(i+1)+" - "+"Варианты добавлены");
				}
				catch(...){
					Application->Title="Ошибка";
					ShowMessage("Ошибка при записи вопроса "+TestQuestions[i].getText());
					WordApplication.OleProcedure("Quit");
					return 1;
				}
			}
			else if (TestQuestions[i].getQuestionType()=="somevariants") {
					FTest->MemoCheck->Lines->Add(IntToStr(i+1)+" - "+"Тип вопроса - с несколькими ответами");
				try{
					int *usedPositions=new int[variantsAmount+1];
					int position;
						FTest->MemoCheck->Lines->Add(IntToStr(i+1)+" - "+"Запись вариантов");
					for (j = 0; j <variantsAmount+1; j++) {
						position=generateRandomNumber(usedPositions,variantsAmount+1,variantsAmount+1);
						usedPositions[j]=position;
							FTest->MemoCheck->Lines->Add(IntToStr(i+1)+" - "+"Позиция: "+IntToStr(position));
						addParagraph(&paragraphsCounter,&Paragraphs,"\t"+TestQuestions[i].getVariant(position-1));
					}
					delete []usedPositions;
						FTest->MemoCheck->Lines->Add(IntToStr(i+1)+" - "+"Варианты расставлены");
				}
				catch(...){
					Application->Title="Ошибка";
					ShowMessage("Ошибка при записи вопроса "+TestQuestions[i].getText());
					WordApplication.OleProcedure("Quit");
					return 1;
				}
			}

			try{
					FTest->MemoCheck->Lines->Add(IntToStr(i+1)+" - "+"Запись правильного и полученного ответов");
				if (TestQuestions[i].getQuestionType()=="choose" || TestQuestions[i].getQuestionType()=="write"){
					addParagraph(&paragraphsCounter,&Paragraphs,"Ответ тестируемого: "+TestQuestions[i].getUserAnswer());
					addParagraph(&paragraphsCounter,&Paragraphs,"Правильный ответ: "+TestQuestions[i].getRightVariant());
				}
				else if (TestQuestions[i].getQuestionType()=="arrange" ){
					addParagraph(&paragraphsCounter,&Paragraphs,"Ответ тестируемого:\n");
					for (j = 0; j <ARRANGES; j++) {
						addParagraph(&paragraphsCounter,&Paragraphs,"\t"+TestQuestions[i].getVariant(j)+" --- "+TestQuestions[i].getUserArrangeAnswer(j));
					}
					addParagraph(&paragraphsCounter,&Paragraphs,"Правильный ответ:\n");
					for (j=0; j<ARRANGES; j++) {
						addParagraph(&paragraphsCounter,&Paragraphs,"\t"+TestQuestions[i].getVariant(j)+" --- "+TestQuestions[i].getVariant(j+ARRANGES));
					}
				}
				else if (TestQuestions[i].getQuestionType()=="arrangepictures") {
					addParagraph(&paragraphsCounter,&Paragraphs,"Ответ тестируемого:\n");
					for (j = 0; j <TestQuestions[i].getRightVariantsAmount(); j++) {
						addParagraph(&paragraphsCounter,&Paragraphs,"\t"+IntToStr(j+1)+". "+TestQuestions[i].getUserArrangeAnswer(j));
					}
					addParagraph(&paragraphsCounter,&Paragraphs,"Правильный ответ:\n");
					for (j=0; j<TestQuestions[i].getRightVariantsAmount(); j++) {
						addParagraph(&paragraphsCounter,&Paragraphs,"\t"+IntToStr(j+1)+". "+TestQuestions[i].getVariant(j));
					}
				}
				else if (TestQuestions[i].getQuestionType()=="somevariants") {
					addParagraph(&paragraphsCounter,&Paragraphs,"Ответ тестируемого: "+TestQuestions[i].getUserAnswer());
					AnsiString answer="";
					for (j = 0; j < TestQuestions[i].getRightVariantsAmount(); j++) {
						answer+=TestQuestions[i].getVariant(j)+", ";
					}
					addParagraph(&paragraphsCounter,&Paragraphs,"Правильный ответ: "+answer);
				}
				else if (TestQuestions[i].getQuestionType()=="video" || TestQuestions[i].getQuestionType()=="music") {
					addParagraph(&paragraphsCounter,&Paragraphs,"Ответ тестируемого: "+TestQuestions[i].getUserAnswer());
					addParagraph(&paragraphsCounter,&Paragraphs,"Правильный ответ: "+TestQuestions[i].getRightVariant());
				}
					FTest->MemoCheck->Lines->Add(IntToStr(i+1)+" - "+"Ответы записаны");
			}
			catch(...){
				Application->Title="Ошибка";
				ShowMessage("Ошибка при записи правильного ответа в вопросе "+TestQuestions[i].getText());
				WordApplication.OleProcedure("Quit");
				return 1;
			}

				FTest->MemoCheck->Lines->Add(IntToStr(i+1)+" - "+"Запись полученных баллов за вопрос");
			try{
				int difficulty=TestQuestions[i].getDifficulty();
				if (TestQuestions[i].getIsRightAnswer()==true) {
					switch (difficulty) {
						case 1: addParagraph(&paragraphsCounter,&Paragraphs,"Ответ засчитан. Получен 1 балл за вопрос"); break;
						case 2:
						case 3:
						case 4: addParagraph(&paragraphsCounter,&Paragraphs,"Ответ засчитан. Получено "+IntToStr(difficulty)+" балла за вопрос");
					}
					//addParagraph(&paragraphsCounter,&Paragraphs,"Ответ засчитан. Получено "+IntToStr(TestQuestions[i].getDifficulty())+" баллов за вопрос");
				}
				else addParagraph(&paragraphsCounter,&Paragraphs,"Ответ не засчитан. Баллы ("+IntToStr(difficulty)+") за вопрос не начислены");
					FTest->MemoCheck->Lines->Add(IntToStr(i+1)+" - "+"Баллы за вопрос записаны");
			}
			catch(...){
				Application->Title="Ошибка";
				ShowMessage("Ошибка при записи баллов за вопрос "+TestQuestions[i].getText());
				WordApplication.OleProcedure("Quit");
				return 1;
			}
			FTest->MemoCheck->Lines->Add(IntToStr(i+1)+" - "+"Запись вопроса закончена");
			FTest->MemoCheck->Lines->Add("");
		}

		FTest->LabelQuestionNumber->Caption="Завершение";
		FTest->LabelQuestionNumber->Left=(FTest->ClientWidth-FTest->LabelQuestionNumber->Width)/2;
		addParagraph(&paragraphsCounter,&Paragraphs,"");
		addParagraph(&paragraphsCounter,&Paragraphs,"");

		FTest->MemoCheck->Lines->Add("Запись результата");
		addParagraph(&paragraphsCounter,&Paragraphs,"Результат: "+IntToStr(testStatistics.getCollectedPoints())+" из " +IntToStr(testStatistics.getMaxPoints())+" баллов набрано");
		addParagraph(&paragraphsCounter,&Paragraphs,"Оценка: " + IntToStr(testStatistics.testResult()));
		FTest->MemoCheck->Lines->Add("Результат записан");

		if (FCharacteristics->CheckBoxAutoCreateDocument->Checked==false) WordApplication.OlePropertySet("Visible",true);
		else{
			Document.OleProcedure("SaveAs",StringToOleStr(fileName));
			WordApplication.OleProcedure("Quit");
			FTest->LabelQuestionNumber->Caption="Завершено";
			FTest->LabelQuestionNumber->Left=(FTest->ClientWidth-FTest->LabelQuestionNumber->Width)/2;
			Application->Title="Отчет создан";
			ShowMessage("Отчет по тестированию успешно создан.");
			FTest->LabelQuestionNumber->Visible=false;
			if (FCharacteristics->CheckBoxAutoCreateDocument->Checked==true) FTest->ButtonGoBack->Left=(FTest->ClientWidth-FTest->ButtonGoBack->Width)/2;
		}
		FTest->MemoCheck->Lines->Add("Отчет создан");
		FTest->MemoCheck->Left=INDENTION;
	}
	catch(...){
		WordApplication.OleProcedure("Quit");
		Application->Title="Ошибка";
		Application->Title="Ошибка";
		ShowMessage("Произошла ошибка при оформлении отчета");
		return 1;
	}

	return 0;
}
//---------------------------------------------------------------------------
void __fastcall TFTest::ButtonCreateDocumentClick(TObject *Sender)
{
	if (ButtonCreateDocument->Tag==0) {
		ButtonCreateDocument->Caption="Создание...";
		ButtonCreateDocument->Tag=1;
		if (creatingDocument()!=0){
			MemoCheck->Lines->SaveToFile(CURRENT_DIRECTORY+"\\report.txt");
			Application->Title="Ошибка";
			ShowMessage("Отчет об ошибке создан. Расположение: "+CURRENT_DIRECTORY+"\\report.txt");
		}
		ButtonCreateDocument->Caption="Отчет оформлен";
		if (isCreatedTestByProgram==false) TimerForButton->Enabled=true;
		else {
			ButtonCreateDocument->Caption="Оформить отчет";
			ButtonCreateDocument->Tag=0;
		}
	}
}
//---------------------------------------------------------------------------
void addToCell(Variant Sheet,int row,int col,AnsiString value){
	Variant Cell;
	Cell=Sheet.OlePropertyGet("Cells").OlePropertyGet("Item",row,col);
	Cell.OlePropertySet("Value",StringToOleStr(value));
}

void saveTest(AnsiString FileName){
		if (FCharacteristics->CheckBoxAutoSaveTest->Checked == true) FTest->LabelQuestionNumber->Visible=true;
		FTest->LabelQuestionNumber->Caption="Подготовка";
		FTest->LabelQuestionNumber->Left=(FTest->ClientWidth-FTest->LabelQuestionNumber->Width)/2;
		int i,j;
		Variant ExcelApplication,ExcelBooks,Book,Sheet,Sheets,Cell;

		try{
			FTest->LabelQuestionNumber->Caption="Открытие Microsoft Excel";
			FTest->LabelQuestionNumber->Left=(FTest->ClientWidth-FTest->LabelQuestionNumber->Width)/2;
			ExcelApplication=CreateOleObject("Excel.Application");
			ExcelBooks=ExcelApplication.OlePropertyGet("Workbooks");
			ExcelBooks.OleProcedure("Add");
			Book=ExcelBooks.OlePropertyGet("Item",1);
			Sheets=Book.OlePropertyGet("Worksheets");
			Sheet=Sheets.OlePropertyGet("Item",1);
			Sheet.OlePropertySet("Name",StringToOleStr("Тестирование "+dateNow()));
		}
		catch(...){
			Application->Title="Ошибка";
			ShowMessage("Ошибка при открытии Excel");
			ExcelApplication.OleProcedure("Quit");
			return;
		}

        FTest->ProgressBar1->Width=FTest->ClientWidth-4*INDENTION;
		FTest->ProgressBar1->Max=questionsAmount;
		FTest->ProgressBar1->Step=1;
		FTest->ProgressBar1->Position=0;
		FTest->ProgressBar1->Visible=true;

		try{
			for (i=0; i < questionsAmount; i++) {
				FTest->LabelQuestionNumber->Caption="Вопрос №"+IntToStr(i+1);
				FTest->LabelQuestionNumber->Left=(FTest->ClientWidth-FTest->LabelQuestionNumber->Width)/2;
				FTest->ProgressBar1->StepIt();
				addToCell(Sheet,i+1,1,TestQuestions[i].getText());  //A
				if (TestQuestions[i].getQuestionType()=="choose" || TestQuestions[i].getQuestionType()=="write" || TestQuestions[i].getQuestionType()=="music" || TestQuestions[i].getQuestionType()=="video") {
					addToCell(Sheet,i+1,2,TestQuestions[i].getRightVariant()); //B
					for (j=0; j < variantsAmount; j++) {
						addToCell(Sheet,i+1,3+j,TestQuestions[i].getVariant(j)); //от C до K
					}
				}
				else if (TestQuestions[i].getQuestionType()=="arrange"){
					for (j=0; j < ARRANGES*2; j++) {
						addToCell(Sheet,i+1,2+j,TestQuestions[i].getVariant(j)); //от B до ...
					}
				}
				else if (TestQuestions[i].getQuestionType()=="arrangepictures"){
					addToCell(Sheet,i+1,2,TestQuestions[i].getRightVariantsAmount());
					for (j=0; j < TestQuestions[i].getRightVariantsAmount(); j++) {
						addToCell(Sheet,i+1,3+j,TestQuestions[i].getVariant(j)); //от C до ...
					}
				}
				else if (TestQuestions[i].getQuestionType()=="somevariants") {
					addToCell(Sheet,i+1,2,TestQuestions[i].getRightVariantsAmount());
					for (j=0; j < variantsAmount+1; j++) {
						addToCell(Sheet,i+1,3+j,TestQuestions[i].getVariant(j)); //от C до ...
					}
				}

				if (TestQuestions[i].getQuestionType()=="music" || TestQuestions[i].getQuestionType()=="video") {
					addToCell(Sheet,i+1,21,TestQuestions[i].getWayToMediaFile());    //U
				}
				if (TestQuestions[i].getQuestionType()=="video") {
					addToCell(Sheet,i+1,24,IntToStr(TestQuestions[i].getVideoResolutionX()));  //X
					addToCell(Sheet,i+1,25,IntToStr(TestQuestions[i].getVideoResolutionY()));  //Y
				}

				addToCell(Sheet,i+1,13,IntToStr(TestQuestions[i].getDifficulty()));//M
				addToCell(Sheet,i+1,1,TestQuestions[i].getText());
				for (j=0; j < IMAGES_AMOUNT; j++) {
					addToCell(Sheet,i+1,14+j,TestQuestions[i].getPicture(j));  //N-Q
				}
				addToCell(Sheet,i+1,20,TestQuestions[i].getQuestionType());//T
				addToCell(Sheet,i+1,23,TestQuestions[i].getReference());//W
			}
			FTest->LabelQuestionNumber->Caption="Завершение";
			FTest->LabelQuestionNumber->Left=(FTest->ClientWidth-FTest->LabelQuestionNumber->Width)/2;
			addToCell(Sheet,1,18,questionsAmount);//R
			addToCell(Sheet,1,19,variantsAmount+1);//S
			ExcelApplication.OlePropertyGet("Workbooks").OlePropertyGet("Item",1).OleProcedure("SaveAs",StringToOleStr(FileName+".xlsx"));
			ExcelApplication.OleProcedure("Quit");
			if (FCharacteristics->CheckBoxAutoSaveTest->Checked==true) {
				FTest->LabelQuestionNumber->Caption="Завершено";
				FTest->LabelQuestionNumber->Left=(FTest->ClientWidth-FTest->LabelQuestionNumber->Width)/2;
				Application->Title="Тестирование сохранено";
				ShowMessage("Тестирование успешно сохранено.");
				FTest->LabelQuestionNumber->Visible=false;
				FTest->ProgressBar1->Visible=false;
			}
		}
		catch(...){
			Application->Title="Ошибка";
			ShowMessage("Ошибка при сохранении тестирования");
			ExcelApplication.OleProcedure("Quit");
		}
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
		info->timeIndex = FCharacteristics->RadioGroupTime->ItemIndex;
		isRight = false;
	}
	if (info->questionsAmount < 1 || info->questionsAmount > 80	) {
		info->questionsAmount = questionsAmount;
		isRight  =false;
	}
	if (info->variantsAmount < 4 || info->variantsAmount > 10) {
		info->variantsAmount = variantsAmount;
		isRight = false;
	}
	return isRight;
}

void __fastcall TFTest::ButtonSaveTestClick(TObject *Sender)
{
	if (ButtonSaveTest->Tag==1) {
		return;
	}

	try{
		AnsiString previousName=EditTestName->Text;
		EditTestName->Text=findFiles(EditTestName->Text,"\\Testings",".xlsx");
		if (EditTestName->Text != previousName) {
			Application->Title="Предупреждение";
			if (mrNo==MessageDlg("Тестирование с указанным именем уже существует.\nИмя тестирования было заменено на "+EditTestName->Text+"\nЖелаете продолжить?",mtWarning,TMsgDlgButtons()<<mbYes<<mbNo,0)){
				EditTestName->Text=previousName;
				return;
			}
		}
	}
	catch(...){}

	try{
		ButtonSaveTest->Caption="Сохранение...";
		ButtonSaveTest->Tag=1;
		AnsiString FileName=CURRENT_DIRECTORY+"\\Testings\\"+EditTestName->Text;
		saveTest(FileName);

		AnsiString directory=FileName+".txt";
		FILE* statisticsFile=fopen(directory.c_str(),"w");
		if (statisticsFile == NULL) {
			Application->Title="Ошибка";
			ShowMessage("Критическая ошибка при сохранении результатов тестирования. Результаты не сохранены");
			return;
		}

		TestInfo info;
		info.amountOfPassings=1;
		info.lastResult=testStatistics.testResult();
		info.isTimeLimited=FCharacteristics->CheckBoxTime->Checked;
		info.timeIndex=FCharacteristics->RadioGroupTime->ItemIndex;
		info.questionsAmount=questionsAmount;
		info.variantsAmount=variantsAmount+1;
		info.lastTime=testStatistics.getTime();
		info.areImages=FCharacteristics->CheckBoxImages->Checked;
		AnsiString name=FStart->ComboBoxUsers->Text + "\n";
		strcpy(info.user,name.c_str());
		AnsiString theme=FCharacteristics->ComboBoxTheme->Text + "\n";
		strcpy(info.theme,theme.c_str());

		checkInfo(&info);

		writeInfoToFile(info,statisticsFile);
		fclose(statisticsFile);

		if (isSavedTestByProgram==false){
			TimerForButton->Enabled=true;
			ButtonSaveTest->Enabled=false;
			EditTestName->Enabled=false;
		}
		else ButtonSaveTest->Tag=0;
	}
	catch(...){
		Application->Title="Ошибка";
		ShowMessage("Произошла ошибка при сохранении результатов тестирования");
	}
}
//---------------------------------------------------------------------------
void upDown(){
	FTest->ButtonUp->Left=panelsAnswers[upDownPosition]->Left+panelsAnswers[upDownPosition]->Width;
	FTest->ButtonDown->Left=panelsAnswers[upDownPosition]->Left+panelsAnswers[upDownPosition]->Width;
	FTest->ButtonUp->Top=panelsAnswers[upDownPosition]->Top;
	FTest->ButtonDown->Top=panelsAnswers[upDownPosition]->Top+FTest->ButtonUp->Height;
	FTest->ButtonUpGeneral->Left=panelsAnswers[upDownPosition]->Left+panelsAnswers[upDownPosition]->Width+7*INDENTION;
	FTest->ButtonDownGeneral->Left=panelsAnswers[upDownPosition]->Left+panelsAnswers[upDownPosition]->Width+7*INDENTION;
	FTest->ButtonUpGeneral->Top=panelsAnswers[0]->Top+(panelsAnswers[ARRANGES-1]->Top+panelsAnswers[ARRANGES-1]->Height-panelsAnswers[0]->Top-2*FTest->ButtonUpGeneral->Height)/2;
	FTest->ButtonDownGeneral->Top=FTest->ButtonUpGeneral->Top+FTest->ButtonUpGeneral->Height;
}
void leftRight(){
	FTest->ButtonLeft->Left=panelsAnswers[upDownPosition]->Left - FTest->ButtonLeft->Width;
	FTest->ButtonLeft->Top=panelsAnswers[upDownPosition]->Top;
	FTest->ButtonRight->Left=panelsAnswers[upDownPosition]->Left + panelsAnswers[upDownPosition]->Width;
	FTest->ButtonRight->Top=panelsAnswers[upDownPosition]->Top;
	FTest->ButtonLeft->Height=panelsAnswers[upDownPosition]->Height;
    FTest->ButtonRight->Height=panelsAnswers[upDownPosition]->Height;
	FTest->ButtonLeftGeneral->Left=(FTest->ClientWidth - 2*(FTest->ButtonLeftGeneral->Width+INDENTION))/2;
	FTest->ButtonRightGeneral->Left=FTest->ButtonLeftGeneral->Left + FTest->ButtonLeftGeneral->Width +2*INDENTION;
	FTest->ButtonLeftGeneral->Top=panelsAnswers[0]->Top+panelsAnswers[0]->Height+5*INDENTION;
	FTest->ButtonRightGeneral->Top=FTest->ButtonLeftGeneral->Top;
}
void changing(int changingPosition){
	int amount;
	if (TestQuestions[globalCounter].getQuestionType()=="arrange") {
		amount=ARRANGES;
	}
	else if (TestQuestions[globalCounter].getQuestionType()=="arrangepictures") {
		amount=TestQuestions[globalCounter].getRightVariantsAmount();
	}
	if (changingPosition<0) {
		changingPosition=amount-1;
	}
	if (changingPosition>=amount) {
		changingPosition=0;
	}
	AnsiString temp;
	temp=panelsAnswers[changingPosition]->Caption;
	panelsAnswers[changingPosition]->Caption=panelsAnswers[upDownPosition]->Caption;
	panelsAnswers[upDownPosition]->Caption=temp;
}
void __fastcall TFTest::ButtonDownGeneralClick(TObject *Sender)
{
	upDownPosition++;
	if (upDownPosition>=ARRANGES) {
		upDownPosition=0;
	}
	upDown();
}
//---------------------------------------------------------------------------
void __fastcall TFTest::ButtonUpGeneralClick(TObject *Sender)
{
	upDownPosition--;
	if (upDownPosition<0) {
		upDownPosition=ARRANGES-1;
	}
	upDown();
}
//---------------------------------------------------------------------------
void __fastcall TFTest::ButtonUpClick(TObject *Sender)
{
	int changingPosition=upDownPosition-1;
	changing(changingPosition);
}
//---------------------------------------------------------------------------
void __fastcall TFTest::ButtonDownClick(TObject *Sender)
{
	int changingPosition=upDownPosition+1;
	changing(changingPosition);
}
//---------------------------------------------------------------------------
void __fastcall TFTest::ButtonLeftGeneralClick(TObject *Sender)
{
	upDownPosition--;
	if (upDownPosition<0) {
		upDownPosition=TestQuestions[globalCounter].getRightVariantsAmount()-1;
	}
	leftRight();
}
//---------------------------------------------------------------------------
void __fastcall TFTest::ButtonRightGeneralClick(TObject *Sender)
{
	upDownPosition++;
	if (upDownPosition>=TestQuestions[globalCounter].getRightVariantsAmount()) {
		upDownPosition=0;
	}
	leftRight();
}
//---------------------------------------------------------------------------
void __fastcall TFTest::MemoQuestionTextClick(TObject *Sender)
{
	UnicodeString temp;
	if (testStatistics.getIsEndOfTest()==false){
		Application->Title="Помощь";
		ShowMessage(MemoQuestionText->Hint);
	}
	else if (questionIndex>=0 && questionIndex!=-1){
		temp=TestQuestions[questionIndex].getReference();
		if (temp!="")
			if (temp[1]=='h') ShellExecute(Handle,L"open",temp.w_str(),NULL,NULL,SW_SHOW);
	}
}
//---------------------------------------------------------------------------
void __fastcall TFTest::MenuArrowClick(TObject *Sender)
{
	MenuArrow->Checked=true;
	MenuDropDown->Checked=false;
}
//---------------------------------------------------------------------------
void __fastcall TFTest::MenuDropDownClick(TObject *Sender)
{
	MenuArrow->Checked=false;
	MenuDropDown->Checked=true;
}
//---------------------------------------------------------------------------
void __fastcall TFTest::StringGridResultsSelectCell(TObject *Sender, int ACol, int ARow,
		  bool &CanSelect)
{
	if (isAllowedToShow == false) {
		return;
	}
	int i;
	StringGridResults->Left=2*INDENTION;
	MemoQuestionText->Left=StringGridResults->Left+StringGridResults->Width+3*INDENTION;
	MemoQuestionText->Top=StringGridResults->Top;
	MemoQuestionText->Width=FTest->ClientWidth-StringGridResults->Width-8*INDENTION;
	MemoQuestionText->Lines->Clear();
	MemoQuestionText->Lines->Add(TestQuestions[ARow-1].getText());
	MemoQuestionText->Font->Height=FONT_HEIGHT*0.8;
	MemoQuestionText->Visible=true;
	MemoQuestionText->ShowHint=false;

	MemoQuestionText->Lines->Add("");
	if (TestQuestions[ARow-1].getQuestionType()=="choose" || TestQuestions[ARow-1].getQuestionType()=="write" || TestQuestions[ARow-1].getQuestionType()=="music" || TestQuestions[ARow-1].getQuestionType()=="video") {
		MemoQuestionText->Lines->Add("Правильный ответ: "+TestQuestions[ARow-1].getRightVariant());
		MemoQuestionText->Lines->Add("Ваш ответ: "+TestQuestions[ARow-1].getUserAnswer());
	}
	else if (TestQuestions[ARow-1].getQuestionType()=="arrange") {
		MemoQuestionText->Lines->Add("Правильный ответ: ");
		for (i = 0; i < ARRANGES; i++) {
			MemoQuestionText->Lines->Add(TestQuestions[ARow-1].getVariant(i)+" - "+TestQuestions[ARow-1].getVariant(i+ARRANGES));
		}
		MemoQuestionText->Lines->Add("");
		MemoQuestionText->Lines->Add("Ваш ответ: ");
		   for (i = 0; i < ARRANGES; i++) {
			MemoQuestionText->Lines->Add(TestQuestions[ARow-1].getVariant(i)+" - "+TestQuestions[ARow-1].getUserArrangeAnswer(i));
		}
	}
	else if (TestQuestions[ARow-1].getQuestionType()=="somevariants"){
		int amount=TestQuestions[ARow-1].getRightVariantsAmount();
		AnsiString tempAnswer="";
		for (i = 0; i < amount; i++) {
			tempAnswer+=TestQuestions[ARow-1].getVariant(i);
			if (i<amount-1) {
				tempAnswer+=", ";
			}
			else tempAnswer+=".";
		}
		MemoQuestionText->Lines->Add("Правильный ответ: "+tempAnswer);
		MemoQuestionText->Lines->Add("Ваш ответ: "+TestQuestions[ARow-1].getUserAnswer());
	}
	else if (TestQuestions[ARow-1].getQuestionType()=="arrangepictures") {
		int amount=TestQuestions[ARow-1].getRightVariantsAmount();
		MemoQuestionText->Lines->Add("Правильный ответ: ");
		for (i = 0; i < amount; i++) {
			MemoQuestionText->Lines->Add(IntToStr(i+1)+". - "+TestQuestions[ARow-1].getVariant(i));
		}
		MemoQuestionText->Lines->Add("");
        MemoQuestionText->Lines->Add("Ваш ответ: ");
		for (i = 0; i < amount; i++) {
			MemoQuestionText->Lines->Add(IntToStr(i+1)+". - "+TestQuestions[ARow-1].getUserArrangeAnswer(i));
		}
	}

	MemoQuestionText->Lines->Add("");
	if (TestQuestions[ARow-1].getReference()!="") {
		MemoQuestionText->Lines->Add("Нажмите, чтобы получить дополнительную информацию по вопросу...");
	}
	questionIndex=ARow-1;

	MemoQuestionText->Height=(FTest->MemoQuestionText->Lines->Count+1)*FONT_HEIGHT*0.8;
	MemoQuestionText->Perform(EM_SCROLL,SB_LINEUP,0);


	ButtonGoBack->Left=MemoQuestionText->Left;
	ButtonCreateDocument->Left=ButtonGoBack->Left+ButtonGoBack->Width+3*INDENTION;
	ButtonGoBack->Top=MemoQuestionText->Top+MemoQuestionText->Height+3*INDENTION;
	ButtonCreateDocument->Top=ButtonGoBack->Top;
	EditTestName->Top=ButtonGoBack->Top+ButtonGoBack->Height+3*INDENTION+ButtonSaveTest->Height-EditTestName->Height;
	EditTestName->Left=ButtonGoBack->Left;
	ButtonSaveTest->Top=ButtonGoBack->Top+ButtonGoBack->Height+3*INDENTION;
	ButtonSaveTest->Left=EditTestName->Left+EditTestName->Width+3*INDENTION;
	MemoQuestionText->Visible = true;
}
//---------------------------------------------------------------------------
void __fastcall TFTest::ButtonPreviousQuestionClick(TObject *Sender)
{
	if (FCharacteristics->CheckBoxTime->Checked==true) return;
	int i;
	bool isNotAnsweredQuestion=false;

	if (questionsAmount-1==testStatistics.getAmountOfAnsweredQuestions()) {
		return;
	}

	for (i = globalCounter-1; i >=0; i--) {
		if (TestQuestions[i].getIsAnswered()==false) {
			TestQuestions[globalCounter].setIsAnswered(false);
			globalCounter=i;
			changeQuestion(globalCounter);
			isNotAnsweredQuestion=true;
			break;
		}
	}

	if (isNotAnsweredQuestion==false) {
		for (i = questionsAmount-1; i >=globalCounter; i--) {
			if (TestQuestions[i].getIsAnswered()==false) {
				TestQuestions[globalCounter].setIsAnswered(false);
				globalCounter=i;
				changeQuestion(globalCounter);
				isNotAnsweredQuestion=true;
				break;
			}
		}
	}

	if (FCharacteristics->CheckBoxTime->Checked==true) {
		TimerTime->Enabled=false;
		PanelTime->Font->Color=clBlack;
		PanelTime->Caption=timeToString(SECONDS_FOR_ANSWER);
		secondsForAnswer=SECONDS_FOR_ANSWER;
		TimerTime->Enabled=true;
	}
}
//---------------------------------------------------------------------------
void __fastcall TFTest::ButtonNextQuestionClick(TObject *Sender)
{
	if (FCharacteristics->CheckBoxTime->Checked==false) changeToNotAnsweredQuestion();
}
//---------------------------------------------------------------------------
void changingVisibilityOfComponents(bool value,TImage* image){
	int i;

	if (TestQuestions[globalCounter].getQuestionType()=="choose") {
		FTest->RadioGroupVariants->Visible=value;
	}
	else if (TestQuestions[globalCounter].getQuestionType()=="write") {
		FTest->EditAnswer->Visible=value;
		FTest->LabelEnterAnswer->Visible=value;
	}
	else if (TestQuestions[globalCounter].getQuestionType()=="arrange") {
		for (i = 0; i < ARRANGES; i++) {
			panelsInfo[i]->Visible=value;
			if (FTest->MenuArrow->Checked==true) {
				panelsAnswers[i]->Visible=value;
			}
			else if (FTest->MenuDropDown->Checked==true) {
				comboBoxAnswers[i]->Visible=value;
			}
		}
	}
	else if (TestQuestions[globalCounter].getQuestionType()=="somevariants") {
		for (i = 0; i<MAX_VARIANTS_AMOUNT; i++) {
			checkBoxes[i]->Visible=value;
		}
	}
	else if (TestQuestions[globalCounter].getQuestionType()=="arrangepictures") {
		for (i = 0; i < TestQuestions[globalCounter].getImagesAmount(); i++) {
			if (FTest->MenuArrow->Checked==true) {
				panelsAnswers[i]->Visible=value;
			}
			else if (FTest->MenuDropDown->Checked==true) {
				comboBoxAnswers[i]->Visible=value;
			}
		}
	}
	else if (TestQuestions[globalCounter].getQuestionType()=="music") {
		FTest->RadioGroupVariants->Visible=value;
		FTest->MediaPlayer1->Visible=value;
		FTest->ScrollBar1->Visible=value;
	}

	for (i = 0; i < TestQuestions[globalCounter].getImagesAmount(); i++) {
		if (images[i]!=image) {
			images[i]->Visible=value;
		}
	}

	FTest->ButtonAnswer->Visible=value;
	FTest->ButtonNextQuestion->Visible=value;
	FTest->ButtonPreviousQuestion->Visible=value;
	FTest->LabelBottomLine->Visible=value;
}

bool wasClickOnImage=false;
int imageLeftPosition;
int imageTopPosition;
TImage* clickedImage;

void __fastcall TFTest::Image1Click(TObject *Sender)
{
	TImage* image=static_cast<TImage*>(Sender);
	clickedImage=image;
	image->Stretch=true;

	if (wasClickOnImage==false) {
		wasClickOnImage=true;
		changingVisibilityOfComponents(false,image);
		imageLeftPosition=image->Left;
		imageTopPosition=image->Top;

		int height=clickedImage->Picture->Height;
		int width=clickedImage->Picture->Width;
		int left=(FTest->ClientWidth - width)/2;

		if (width>MAX_IMAGE_WIDTH) {
			width=MAX_IMAGE_WIDTH;
			left=(FTest->ClientWidth - width)/2;
		}
		if (height>MAX_IMAGE_HEIGHT) height=MAX_IMAGE_HEIGHT;

		image->Hint="Нажмите, чтобы уменьшить изображение";     //Utf8ToAnsi();
		image->Height=height;
		image->Width=width;
		image->Left=left;
	}
	else {
		wasClickOnImage=false;
		image->Hint="Нажмите, чтобы увеличить изображение";
		image->Height=IMAGE_HEIGHT;
		image->Width=IMAGE_WIDTH;
		image->Left=imageLeftPosition;
		image->Top=imageTopPosition;
		changingVisibilityOfComponents(true,image);
	}
}
//-----------------------------------------------------------------------------
void __fastcall TFTest::TimerForMusicTimer(TObject *Sender)
{
	ScrollBar1->Position++;

	if (ScrollBar1->Position==ScrollBar1->Max) {
		ScrollBar1->Position=0;
		MediaPlayer1->Stop();
		MediaPlayer1->Position=0;
		TimerForMusic->Enabled=false;
	}
}
//---------------------------------------------------------------------------
void __fastcall TFTest::MediaPlayer1Click(TObject *Sender, TMPBtnType Button, bool &DoDefault)
{
	try{
		if (Button==btPlay) {
			if (MediaPlayer1->Position==0) ScrollBar1->Position=0;
			TimerForMusic->Enabled=true;
		}
		else if (Button==btPause) {
			TimerForMusic->Enabled=false;
		}
	}
	catch(...){
		Application->Title="Ошибка";
		ShowMessage("Ошибка при открытии мультимедийного файла");
	}
}
//---------------------------------------------------------------------------
void __fastcall TFTest::ScrollBar1Scroll(TObject *Sender, TScrollCode ScrollCode,
          int &ScrollPos)
{
	MediaPlayer1->Position=ScrollBar1->Position*1000;
	MediaPlayer1->Play();
}
//---------------------------------------------------------------------------
AnsiString findFiles(AnsiString name, AnsiString folder, AnsiString fileType){
	struct _finddata_t fileNames;
	intptr_t file;

	AnsiString directory=CURRENT_DIRECTORY+folder;
	AnsiString searchedName=name+fileType;
	chdir(directory.c_str());
	bool isFoundSimilar=true;
	bool isFirstTime=true;
	bool wasChanged=false;
	int i=0;

	while (isFoundSimilar==true){
		if (isFirstTime==false){
			searchedName=name+"("+IntToStr(i)+")"+fileType;
			wasChanged=true;
		}
		else isFirstTime=false;

		AnsiString format="*"+fileType;
		if ((file=_findfirst(format.c_str(),&fileNames))!=-1) {
			bool isDifferent=true;
			do{
				if (strcmp(searchedName.c_str(),fileNames.name)==0) {
					isDifferent=false;
					break;
				}
			}   while( _findnext(file,&fileNames)==0) ;
			if (isDifferent==false) {
				isFoundSimilar=true;
			}
			else isFoundSimilar=false;
		}
		else {
			isFoundSimilar=false;
		}
		i++;
	}
	chdir(CURRENT_DIRECTORY.c_str());

	if (wasChanged==true) return name+"("+IntToStr(i-1)+")";
	else return name;
}

void __fastcall TFTest::EditTestNameChange(TObject *Sender)
{
	ButtonSaveTest->Caption="Сохранить тестирование как "+EditTestName->Text;
}
//---------------------------------------------------------------------------
int FontHeight;
void __fastcall TFTest::TimerForButtonTimer(TObject *Sender)
{
	static int i=0;
	static int j=0;
	if (ButtonSaveTest->Tag==1) {
		ButtonSaveTest->Font->Color=clGreen;
		FontHeight= ButtonSaveTest->Font->Height;
		ButtonSaveTest->Font->Height=FONT_HEIGHT;
		ButtonSaveTest->Caption="Успешно сохранено";
		i++;
		if (i==2){
			ButtonSaveTest->Tag=0;
			i=0;
			ButtonSaveTest->Font->Height=FontHeight;
			ButtonSaveTest->Font->Color=clBlack;
			FTest->ProgressBar1->Visible=false;
			TimerForButton->Enabled=false;
		}
	}
	else if (ButtonCreateDocument->Tag==1) {
		ButtonCreateDocument->Font->Color=clGreen;
		j++;
		if (j==2){
			ButtonCreateDocument->Tag=0;
			j=0;
			ButtonCreateDocument->Caption="Оформить отчет";
			ButtonCreateDocument->Font->Color=clBlack;
			FTest->ProgressBar1->Visible=false;
			TimerForButton->Enabled=false;
		}
	}
}
//---------------------------------------------------------------------------
void __fastcall TFTest::ButtonStopTestClick(TObject *Sender)
{
	if (wasAnyQuestionAnswered==true && FCharacteristics->CheckBoxTime->Checked==false) {
		Application->Title="Предупреждение";
		if (mrNo==MessageDlg("Внимание!\nНеотвеченные вопросы будут не засчитаны.\nРезультат тестирования будет записан.\nЖелаете продолжить?",mtWarning,TMsgDlgButtons()<<mbYes<<mbNo,0)){
			return;
		}
	}
	int i,j;
	for (i = 0; i < questionsAmount; i++) {     ///
		int imagesAmount=0;
		for (j = 0; j < IMAGES_AMOUNT; j++) {
			if (TestQuestions[i].getPicture(j)!="-") {
				imagesAmount++;
			}
			else break;
		}
		TestQuestions[i].setImagesAmount(imagesAmount);

		if (TestQuestions[i].getIsAnswered()==false) {
			TestQuestions[i].setIsRightAnswer(false);
			TestQuestions[i].setIsAnswered(true);
		}
	}
	if (FCharacteristics->CheckBoxTime->Checked==true) {
		seconds+=SECONDS_FOR_ANSWER-secondsForAnswer;
	}
	stopTest();
}
//---------------------------------------------------------------------------
AnsiString unitOfTime(int time){
	if (time<10) {
		return "0"+IntToStr(time);
	}
	return IntToStr(time);
}
//---------------------------------------------------------------------------
AnsiString timeToString(int time){
	int minute=time/60;
	int second=time-(time/60)*60;
	AnsiString Minute=unitOfTime(minute),Second=unitOfTime(second);
	return Minute+" : "+Second;
}
//---------------------------------------------------------------------------
void __fastcall TFTest::TimerTimeTimer(TObject *Sender)
{
	if (FCharacteristics->CheckBoxTime->Checked==false){
		seconds++;
		PanelTime->Caption=timeToString(seconds);
	}
	else{
		secondsForAnswer--;
		PanelTime->Caption=timeToString(secondsForAnswer);
		if (secondsForAnswer<=5) {
			PanelTime->Font->Color=clRed;
			UnicodeString file=CURRENT_DIRECTORY+"\\resourses\\"+"time.wav";
			PlaySound(file.w_str(),0,SND_ASYNC);
		}
		if (secondsForAnswer==0) {
			if (FCharacteristics->CheckBoxTime->Checked ==false) secondsForAnswer=SECONDS_FOR_ANSWER;
			ButtonStopTest->Click();///////////////////////////////////////////////////
			if (FCharacteristics->CheckBoxTime->Checked == true) {
				secondsForAnswer=SECONDS_FOR_ANSWER;
			}
			LabelEnd->Caption="Время истекло. Тестирование окончено.";
			LabelEnd->Left=(FTest->ClientWidth-LabelEnd->Width)/2;
			TimerTime->Enabled=false;
		}
    }
}
//---------------------------------------------------------------------------
void __fastcall TFTest::FormClick(TObject *Sender)
{
	ShowMessage(MemoQuestionText->Hint);
}
//---------------------------------------------------------------------------
void __fastcall TFTest::FormResize(TObject *Sender)
{
	Bevel1->Width=FTest->ClientWidth;
	Bevel1->Height=FTest->ClientHeight;
}
//---------------------------------------------------------------------------

void __fastcall TFTest::TimerForWaitingTimer(TObject *Sender)
{
	for (int i = 0; i < ARRANGES; i++) {
		comboBoxAnswers[i]->Color=clBtnFace;
	}
	ButtonAnswer->Caption="Ответить";
	TimerForWaiting->Enabled=false;
}
//---------------------------------------------------------------------------


