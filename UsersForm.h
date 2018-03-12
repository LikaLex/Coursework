//---------------------------------------------------------------------------

#ifndef UsersFormH
#define UsersFormH
//---------------------------------------------------------------------------
#include <System.Classes.hpp>
#include <Vcl.Controls.hpp>
#include <Vcl.StdCtrls.hpp>
#include <Vcl.Forms.hpp>
#include <Vcl.ExtCtrls.hpp>
#include <Vcl.Imaging.jpeg.hpp>
//---------------------------------------------------------------------------
class TUsersF : public TForm
{
__published:	// IDE-managed Components
	TLabel *WelkomeLabel;
	TButton *LastResultButton;
	TButton *AllTestsButton;
	TButton *AverageScoreButton;
	TButton *InformationButton;
	TButton *BackToTestButton;
	TButton *LastTimeButton;
	TButton *AllTimeButton;
	void __fastcall LastResultButtonClick(TObject *Sender);
	void __fastcall BackToTestButtonClick(TObject *Sender);
	void __fastcall AllTestsButtonClick(TObject *Sender);
	void __fastcall AverageScoreButtonClick(TObject *Sender);
	void __fastcall InformationButtonClick(TObject *Sender);
	void __fastcall FormActivate(TObject *Sender);
	void __fastcall FormClose(TObject *Sender, TCloseAction &Action);
	void __fastcall LastTimeButtonClick(TObject *Sender);
	void __fastcall AllTimeButtonClick(TObject *Sender);
private:	// User declarations
public:		// User declarations
	__fastcall TUsersF(TComponent* Owner);
};
//---------------------------------------------------------------------------
extern PACKAGE TUsersF *UsersF;
//---------------------------------------------------------------------------
#endif
