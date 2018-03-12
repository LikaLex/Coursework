//---------------------------------------------------------------------------

#ifndef UnitUserH
#define UnitUserH
//---------------------------------------------------------------------------
#endif
#include <vcl.h>
class User{
	private:
		AnsiString name;
		AnsiString surname;
		AnsiString group;
		int lastMark;
		int allTestsResult;
		float averageScore;
		int lastTime;
		int allTime;
	public:
		User();

		int getLastMark();
		void setLastMark(int mark);

		AnsiString getName();
		void setName(AnsiString Name);

		AnsiString getSurname();
		void setSurname(AnsiString Surname);

		AnsiString getGroup();
		void setGroup(AnsiString Group);

		int getAllTestsResult();
		void setAllTestsResult(int result);

		float getAverageScore();
		void setAverageScore(float score);

		int getLastTime();
		void setLastTime(int lasttime);

		int getAllTime();
		void setAllTime(int alltime);

};
