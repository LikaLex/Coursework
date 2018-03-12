//---------------------------------------------------------------------------

#ifndef FormReportsH
#define FormReportsH
//---------------------------------------------------------------------------
#include <System.Classes.hpp>
#include <Vcl.Controls.hpp>
#include <Vcl.StdCtrls.hpp>
#include <Vcl.Forms.hpp>
#include <Vcl.ComCtrls.hpp>
//---------------------------------------------------------------------------
class TFReports : public TForm
{
__published:	// IDE-managed Components
	TTreeView *TreeViewReports;
	TButton *ButtonShow;
	TButton *ButtonDelete;
	TButton *ButtonDeleteAll;
	TButton *ButtonBack;
	void __fastcall FormActivate(TObject *Sender);
	void __fastcall FormClose(TObject *Sender, TCloseAction &Action);
	void __fastcall ButtonShowClick(TObject *Sender);
	void __fastcall ButtonDeleteClick(TObject *Sender);
	void __fastcall ButtonDeleteAllClick(TObject *Sender);
	void __fastcall ButtonBackClick(TObject *Sender);
private:	// User declarations
public:		// User declarations
	__fastcall TFReports(TComponent* Owner);
};
//---------------------------------------------------------------------------
extern PACKAGE TFReports *FReports;
//---------------------------------------------------------------------------
#endif
