//---------------------------------------------------------------------------

#ifndef TestsRatingFormH
#define TestsRatingFormH
//---------------------------------------------------------------------------
#include <System.Classes.hpp>
#include <Vcl.Controls.hpp>
#include <Vcl.StdCtrls.hpp>
#include <Vcl.Forms.hpp>
#include <Vcl.Grids.hpp>
//---------------------------------------------------------------------------
class TFormTestsRating : public TForm
{
__published:	// IDE-managed Components
	TButton *BackButton;
	TButton *ButtonAll;
	TButton *ButtonOld;
	TButton *ButtonVKL;
	TButton *ButtonRP;
	TButton *ButtonRI;
	TButton *ButtonTwenty;
	TButton *ButtonNoWar;
	TButton *ButtonWAR;
	TButton *ButtonAfterWar;
	TButton *ButtonMedia;
	TComboBox *ComboBoxChoise;
	TStringGrid *StringGrid1;
	void __fastcall BackButtonClick(TObject *Sender);
	void __fastcall FormClose(TObject *Sender, TCloseAction &Action);
	void __fastcall ButtonAllClick(TObject *Sender);
	void __fastcall ButtonOldClick(TObject *Sender);
	void __fastcall ButtonRPClick(TObject *Sender);
	void __fastcall ButtonRIClick(TObject *Sender);
	void __fastcall ButtonVKLClick(TObject *Sender);
	void __fastcall ButtonNoWarClick(TObject *Sender);
	void __fastcall ButtonAfterWarClick(TObject *Sender);
	void __fastcall ButtonMediaClick(TObject *Sender);
	void __fastcall ButtonTwentyClick(TObject *Sender);
	void __fastcall ButtonWARClick(TObject *Sender);
	void __fastcall ComboBoxChoiseChange(TObject *Sender);
	void __fastcall FormActivate(TObject *Sender);
	void __fastcall ScrollBar1Scroll(TObject *Sender, TScrollCode ScrollCode, int &ScrollPos);




private:	// User declarations
public:		// User declarations
	__fastcall TFormTestsRating(TComponent* Owner);
};
//---------------------------------------------------------------------------
extern PACKAGE TFormTestsRating *FormTestsRating;
//---------------------------------------------------------------------------
#endif
