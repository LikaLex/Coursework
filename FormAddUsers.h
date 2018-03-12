//---------------------------------------------------------------------------

#ifndef FormAddUsersH
#define FormAddUsersH
//---------------------------------------------------------------------------
#include <System.Classes.hpp>
#include <Vcl.Controls.hpp>
#include <Vcl.StdCtrls.hpp>
#include <Vcl.Forms.hpp>
#include <Vcl.ExtCtrls.hpp>
#include <Vcl.Imaging.jpeg.hpp>
#include <Vcl.Imaging.GIFImg.hpp>
//---------------------------------------------------------------------------
class TFAddUser : public TForm
{
__published:	// IDE-managed Components
	TEdit *NameBox;
	TEdit *SurnameBox;
	TEdit *GroupBox;
	TButton *AddButton;
	TLabel *NameLabel;
	TLabel *SurnameLabel;
	TLabel *GroupLabel;
	TButton *ReturnButton;
	TImage *WelkomeImage;
	TMemo *MemoHelp;
	void __fastcall AddButtonClick(TObject *Sender);
	void __fastcall NameBoxChange(TObject *Sender);
	void __fastcall SurnameBoxChange(TObject *Sender);
	void __fastcall GroupBoxChange(TObject *Sender);
	void __fastcall ReturnButtonClick(TObject *Sender);
	void __fastcall FormClose(TObject *Sender, TCloseAction &Action);
	void __fastcall NameBoxClick(TObject *Sender);
	void __fastcall SurnameBoxClick(TObject *Sender);
	void __fastcall NameBoxExit(TObject *Sender);
	void __fastcall SurnameBoxExit(TObject *Sender);
	void __fastcall FormActivate(TObject *Sender);
private:	// User declarations
public:		// User declarations
	__fastcall TFAddUser(TComponent* Owner);
};
//---------------------------------------------------------------------------
extern PACKAGE TFAddUser *FAddUser;
//---------------------------------------------------------------------------
#endif
