//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop
#include <tchar.h>
//---------------------------------------------------------------------------
USEFORM("TestsRatingForm.cpp", FormTestsRating);
USEFORM("RecordsForm.cpp", FormRecords);
USEFORM("RatingForm.cpp", FormRating);
USEFORM("UsersForm.cpp", UsersF);
USEFORM("FormTest.cpp", FTest);
USEFORM("FormAddUsers.cpp", FAddUser);
USEFORM("FormStart.cpp", FStart);
USEFORM("FormReports.cpp", FReports);
USEFORM("FormCharacteristics.cpp", FCharacteristics);
//---------------------------------------------------------------------------
int WINAPI _tWinMain(HINSTANCE, HINSTANCE, LPTSTR, int)
{
	try
	{
		Application->Initialize();
		Application->MainFormOnTaskBar = true;
		Application->CreateForm(__classid(TFStart), &FStart);
		Application->CreateForm(__classid(TFAddUser), &FAddUser);
		Application->CreateForm(__classid(TFCharacteristics), &FCharacteristics);
		Application->CreateForm(__classid(TFReports), &FReports);
		Application->CreateForm(__classid(TFTest), &FTest);
		Application->CreateForm(__classid(TFormRating), &FormRating);
		Application->CreateForm(__classid(TFormRecords), &FormRecords);
		Application->CreateForm(__classid(TFormTestsRating), &FormTestsRating);
		Application->CreateForm(__classid(TUsersF), &UsersF);
		Application->Run();
	}
	catch (Exception &exception)
	{
		Application->ShowException(&exception);
	}
	catch (...)
	{
		try
		{
			throw Exception("");
		}
		catch (Exception &exception)
		{
			Application->ShowException(&exception);
		}
	}
	return 0;
}
//---------------------------------------------------------------------------
