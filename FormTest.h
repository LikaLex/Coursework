//---------------------------------------------------------------------------

#ifndef FormTestH
#define FormTestH
//---------------------------------------------------------------------------
#include <System.Classes.hpp>
#include <Vcl.Controls.hpp>
#include <Vcl.StdCtrls.hpp>
#include <Vcl.Forms.hpp>
#include <Vcl.ExtCtrls.hpp>
#include <Vcl.ComCtrls.hpp>
#include <Vcl.Dialogs.hpp>
#include <Vcl.Imaging.pngimage.hpp>
#include <Vcl.Buttons.hpp>
#include <Vcl.Menus.hpp>
#include <Vcl.Imaging.jpeg.hpp>
#include <Vcl.Grids.hpp>
#include <Vcl.MPlayer.hpp>

#include "TestInfo.h"
//---------------------------------------------------------------------------

class TFTest : public TForm
{
__published:	// IDE-managed Components
	TRadioGroup *RadioGroupVariants;
	TMemo *MemoQuestionText;
	TButton *ButtonAnswer;
	TLabel *LabelEnd;
	TButton *ButtonGoBack;
	TLabel *LabelResult;
	TProgressBar *ProgressBar1;
	TButton *ButtonCreateDocument;
	TEdit *EditAnswer;
	TLabel *LabelEnterAnswer;
	TLabel *LabelMark;
	TStringGrid *StringGridResults;
	TBitBtn *ButtonPreviousQuestion;
	TBitBtn *ButtonNextQuestion;
	TLabel *LabelBottomLine;
	TImage *Image1;
	TMediaPlayer *MediaPlayer1;
	TTimer *TimerForMusic;
	TScrollBar *ScrollBar1;
	TPanel *PanelVideo;
	TTimer *TimerForButton;
	TBitBtn *ButtonSaveTest;
	TBitBtn *ButtonLeft;
	TBitBtn *ButtonRight;
	TBitBtn *ButtonDown;
	TBitBtn *ButtonUp;
	TBitBtn *ButtonUpGeneral;
	TBitBtn *ButtonDownGeneral;
	TBitBtn *ButtonRightGeneral;
	TBitBtn *ButtonLeftGeneral;
	TPopupMenu *PopupMenu1;
	TMenuItem *MenuArrow;
	TMenuItem *MenuDropDown;
	TBitBtn *ButtonStopTest;
	TTimer *TimerTime;
	TPanel *PanelTime;
	TLabeledEdit *EditTestName;
	TBevel *Bevel1;
	TTimer *TimerForWaiting;
	TLabel *LabelSaving;
	TLabel *LabelQuestionNumber;
	TMemo *MemoCheck;
	void __fastcall FormActivate(TObject *Sender);
	void __fastcall FormClose(TObject *Sender, TCloseAction &Action);
	void __fastcall FormCreate(TObject *Sender);
	void __fastcall saveResults();
	void __fastcall stopTest();
	void __fastcall ButtonAnswerClick(TObject *Sender);
	void __fastcall ButtonGoBackClick(TObject *Sender);
	void __fastcall ButtonCreateDocumentClick(TObject *Sender);
	void __fastcall ButtonSaveTestClick(TObject *Sender);
	void __fastcall ButtonDownGeneralClick(TObject *Sender);
	void __fastcall ButtonUpGeneralClick(TObject *Sender);
	void __fastcall ButtonUpClick(TObject *Sender);
	void __fastcall ButtonDownClick(TObject *Sender);
	void __fastcall ButtonLeftGeneralClick(TObject *Sender);
	void __fastcall ButtonRightGeneralClick(TObject *Sender);
	void __fastcall MemoQuestionTextClick(TObject *Sender);
	void __fastcall MenuArrowClick(TObject *Sender);
	void __fastcall MenuDropDownClick(TObject *Sender);
	void __fastcall StringGridResultsSelectCell(TObject *Sender, int ACol, int ARow,
		  bool &CanSelect);
	void __fastcall ButtonPreviousQuestionClick(TObject *Sender);
	void __fastcall ButtonNextQuestionClick(TObject *Sender);
	void __fastcall Image1Click(TObject *Sender);
	void __fastcall TimerForMusicTimer(TObject *Sender);
	void __fastcall MediaPlayer1Click(TObject *Sender, TMPBtnType Button, bool &DoDefault);
	void __fastcall ScrollBar1Scroll(TObject *Sender, TScrollCode ScrollCode, int &ScrollPos);
	void __fastcall EditTestNameChange(TObject *Sender);
	void __fastcall TimerForButtonTimer(TObject *Sender);
	void __fastcall ButtonStopTestClick(TObject *Sender);
	void __fastcall TimerTimeTimer(TObject *Sender);
	void __fastcall FormClick(TObject *Sender);
	void __fastcall FormResize(TObject *Sender);
	void __fastcall TimerForWaitingTimer(TObject *Sender);
private:	// User declarations
public:		// User declarations
	__fastcall TFTest(TComponent* Owner);
};

int generateRandomNumber(int* arrayOfUsedElements, int arraySize, int range);
void zeroingArray(int* array,int size);
void setVariants(Question* TestQuestions,int i, int* usedVariants, int question, Variant Sheet);
void makingTest();
bool changeToNotAnsweredQuestion();
void changeQuestion(int index);
void deleteTest();
AnsiString dateNow();
int creatingDocument();
void addParagraph(int* paragraphCounter,Variant Paragraphs,AnsiString text);
void saveTest(AnsiString FileName);
void addToCell(Variant Sheet,int row,int col,AnsiString value);
void upDown();
void leftRight();
void changing(int changingPosition);
void changingVisibilityOfComponents(bool value,TImage* image);
AnsiString findFiles(AnsiString name, AnsiString folder, AnsiString fileType);
AnsiString unitOfTime(int time);
AnsiString convertUnit(int value, AnsiString str);
AnsiString convertTime();
AnsiString timeToString(int time);
//---------------------------------------------------------------------------
extern PACKAGE TFTest *FTest;
//---------------------------------------------------------------------------
#endif
