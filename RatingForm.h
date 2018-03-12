//---------------------------------------------------------------------------

#ifndef RatingFormH
#define RatingFormH
//---------------------------------------------------------------------------
#include <System.Classes.hpp>
#include <Vcl.Controls.hpp>
#include <Vcl.StdCtrls.hpp>
#include <Vcl.Forms.hpp>
#include <Vcl.Grids.hpp>
//---------------------------------------------------------------------------
class TFormRating : public TForm
{
__published:	// IDE-managed Components
	TButton *BackButton;
	TButton *RatingLastTestButton;
	TButton *RatingQantityButton;
	TButton *RatingAllTestsButton;
	TLabel *RatingLastTestLabel;
	TLabel *RatingAllTestsLabel;
	TLabel *RatingQantityLabel;
	TButton *LastTimeButton;
	TButton *AllTimeButton;
	TStringGrid *StringGrid1;
	void __fastcall BackButtonClick(TObject *Sender);
	void __fastcall FormClose(TObject *Sender, TCloseAction &Action);
	void __fastcall RatingLastTestButtonClick(TObject *Sender);
	void __fastcall RatingAllTestsButtonClick(TObject *Sender);
	void __fastcall RatingQantityButtonClick(TObject *Sender);
	void __fastcall FormActivate(TObject *Sender);
	void __fastcall LastTimeButtonClick(TObject *Sender);
	void __fastcall AllTimeButtonClick(TObject *Sender);


private:	// User declarations
public:		// User declarations
	__fastcall TFormRating(TComponent* Owner);
};
//---------------------------------------------------------------------------
extern PACKAGE TFormRating *FormRating;
//---------------------------------------------------------------------------
#endif
