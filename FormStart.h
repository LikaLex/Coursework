//---------------------------------------------------------------------------

#ifndef FormStartH
#define FormStartH
//---------------------------------------------------------------------------
#include <System.Classes.hpp>
#include <Vcl.Controls.hpp>
#include <Vcl.StdCtrls.hpp>
#include <Vcl.Forms.hpp>
#include <Vcl.ExtCtrls.hpp>
#include <Vcl.Dialogs.hpp>
#include "UsersForm.h"
#include <Vcl.Imaging.jpeg.hpp>
#include <Vcl.MPlayer.hpp>
#include <Vcl.ComCtrls.hpp>
#include <FireDAC.Comp.UI.hpp>
#include <FireDAC.Stan.Intf.hpp>
#include <FireDAC.UI.Intf.hpp>
#include <FireDAC.VCLUI.Login.hpp>

#include "TestInfo.h"
//---------------------------------------------------------------------------
class TFStart : public TForm
{
__published:	// IDE-managed Components
	TButton *ButtonStartRandomTest;
	TButton *ButtonSettings;
	TButton *ButtonStartCreatedTest;
	TComboBox *ComboBoxUsers;
	TLabel *WelkomeLabel;
	TButton *AddButton;
	TEdit *NameBox;
	TButton *OKButton;
	TButton *ButtonOfUsers;
	TButton *TableOfRecordsButton;
	TImage *StartImage;
	TTreeView *TreeViewTestings;
	TBevel *BevelFrame;
	TLabel *LabelTestings;
	TMemo *MemoStatistics;
	TButton *ButtonStart;
	TButton *ButtonDeleteTest;
	TTimer *TimerForHiding;
	TTimer *TimerForOpening;
	TButton *ButtonDelAll;
	TLabel *LabelHint;
	void __fastcall ButtonStartRandomTestClick(TObject *Sender);
	void __fastcall ButtonSettingsClick(TObject *Sender);
	void __fastcall ButtonStartCreatedTestClick(TObject *Sender);
	void __fastcall ProgramStart();
	void __fastcall Continue();
	void __fastcall FormActivate(TObject *Sender);
	void __fastcall ComboBoxUsersChange(TObject *Sender);
	void __fastcall AddButtonClick(TObject *Sender);
	void __fastcall OKButtonClick(TObject *Sender);
	void __fastcall NameBoxClick(TObject *Sender);
	void __fastcall ButtonOfUsersClick(TObject *Sender);
	void __fastcall TableOfRecordsButtonClick(TObject *Sender);
	void __fastcall TreeViewTestingsClick(TObject *Sender);
	void __fastcall ButtonStartClick(TObject *Sender);
	void __fastcall FormCreate(TObject *Sender);
	void __fastcall ButtonDeleteTestClick(TObject *Sender);
	void __fastcall TimerForHidingTimer(TObject *Sender);
	void __fastcall TimerForOpeningTimer(TObject *Sender);
	void __fastcall LabelTestingsClick(TObject *Sender);
	void __fastcall ButtonDelAllClick(TObject *Sender);
	void __fastcall ComboBoxUsersSelect(TObject *Sender);

private:	// User declarations
public:		// User declarations
	__fastcall TFStart(TComponent* Owner);
};
void screenParametres();
void getFiles(AnsiString SubDirectory, TTreeView* TreeView, AnsiString Extension, bool* wasAdded);
void open();
int getBackgroundsAmount();
void startChecking();
//---------------------------------------------------------------------------
extern PACKAGE TFStart *FStart;
//---------------------------------------------------------------------------
#endif
