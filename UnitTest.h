//---------------------------------------------------------------------------

#ifndef UnitTestH
#define UnitTestH
//---------------------------------------------------------------------------
#endif
#include <vcl.h>
class Test{
	private:
		int maxPoints;
		int collectedPoints;
		AnsiString wrongAnswers;
		int questionsAmount;
		bool isEndOfTest;
		int amountOfAnsweredQuestions;
		int time;

	public:
		Test();
		void clearing();

		int getMaxPoints();
		void addMaxPoints(int points);
		int getTime();
		void setTime(int value);
		int getCollectedPoints();
		void addCollectedPoints(int points);
		AnsiString getWrongAnswers();
		void addWrongAnswers(AnsiString answer);
		int testResult();
		int getQuestionsAmount();
		void addQuestionsAmount();
		bool getIsEndOfTest();
		void setIsEndOfTest(bool value);
		int getAmountOfAnsweredQuestions();
		void addAmountOfAnsweredQuestions();
};