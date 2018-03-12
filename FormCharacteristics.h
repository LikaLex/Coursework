//---------------------------------------------------------------------------

#ifndef FormCharacteristicsH
#define FormCharacteristicsH
//---------------------------------------------------------------------------
#include <System.Classes.hpp>
#include <Vcl.Controls.hpp>
#include <Vcl.StdCtrls.hpp>
#include <Vcl.Forms.hpp>
#include <Vcl.ComCtrls.hpp>
#include <Vcl.ExtCtrls.hpp>
#include <FireDAC.Comp.UI.hpp>
#include <FireDAC.Stan.Intf.hpp>
#include <FireDAC.UI.Intf.hpp>
#include <FireDAC.VCLUI.Login.hpp>
//---------------------------------------------------------------------------
struct Settings{
	int questionsAmount;
	int variantsAmount;
	int themeIndex;
	bool isTimeLimited;
	int timeIndex;
	bool autoSaveTest;
	bool autoCreateDocument;
	bool areImages;
	bool areMedia;
};

class TFCharacteristics : public TForm
{
__published:	// IDE-managed Components
	TLabel *Label1;
	TEdit *EditQuestionsAmount;
	TLabel *Label2;
	TEdit *EditVariantsAmount;
	TButton *ButtonGoBack;
	TLabel *LabelTheme;
	TComboBox *ComboBoxTheme;
	TUpDown *UpDown1;
	TUpDown *UpDown2;
	TRadioGroup *RadioGroupArrowsOrDropDowns;
	TCheckBox *CheckBoxTime;
	TBevel *Bevel1;
	TBevel *Bevel2;
	TBevel *Bevel3;
	TRadioGroup *RadioGroupTime;
	TBevel *Bevel4;
	TBevel *Bevel5;
	TBevel *Bevel6;
	TCheckBox *CheckBoxAutoSaveTest;
	TCheckBox *CheckBoxAutoCreateDocument;
	TButton *ButtonReports;
	TCheckBox *CheckBoxImages;
	TCheckBox *CheckBoxMedia;
	TBevel *Bevel7;
	void __fastcall EditQuestionsAmountClick(TObject *Sender);
	void __fastcall EditQuestionsAmountEnter(TObject *Sender);
	void __fastcall EditQuestionsAmountExit(TObject *Sender);
	void __fastcall EditVariantsAmountClick(TObject *Sender);
	void __fastcall EditVariantsAmountEnter(TObject *Sender);
	void __fastcall EditVariantsAmountExit(TObject *Sender);
	void __fastcall ButtonGoBackClick(TObject *Sender);
	void __fastcall FormActivate(TObject *Sender);
	void __fastcall FormHide(TObject *Sender);
	void __fastcall FormClose(TObject *Sender, TCloseAction &Action);
	void __fastcall ComboBoxThemeChange(TObject *Sender);
	void __fastcall CheckBoxTimeClick(TObject *Sender);
	void __fastcall ButtonReportsClick(TObject *Sender);
private:	// User declarations
public:		// User declarations
	__fastcall TFCharacteristics(TComponent* Owner);
};

bool isRightData(AnsiString data,int maxValue);
void saveSettings();
void getSettings();
//---------------------------------------------------------------------------
extern PACKAGE TFCharacteristics *FCharacteristics;
//---------------------------------------------------------------------------
#endif
