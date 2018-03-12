//---------------------------------------------------------------------------

#pragma hdrstop

#include "UnitQuestion.h"
#include "FormStart.h"
#include "FormCharacteristics.h"
#include "FormTest.h"
#include "UnitTest.h"
//---------------------------------------------------------------------------
#pragma package(smart_init)
Test::Test(){
}

void Test::clearing(){
	maxPoints=0;
	collectedPoints=0;
	wrongAnswers="";
	questionsAmount;
	isEndOfTest=false;
	amountOfAnsweredQuestions=0;
}

int		   Test::getMaxPoints() 				   { return maxPoints;}
void 	   Test::addMaxPoints(int points) 		   { maxPoints+=points;}
int 	   Test::getTime()                         { return time;}
void	   Test::setTime(int value)                { time=value;}
int 	   Test::getCollectedPoints()			   { return collectedPoints;}
void 	   Test::addCollectedPoints(int points)    { collectedPoints+=points;}
AnsiString Test::getWrongAnswers()				   { return wrongAnswers;}
void 	   Test::addWrongAnswers(AnsiString answer){ wrongAnswers+=answer+", "; }
int 	   Test::getQuestionsAmount()			   { return questionsAmount;}
void 	   Test::addQuestionsAmount()			   { questionsAmount++;}
bool	   Test::getIsEndOfTest()				   { return isEndOfTest;}
void 	   Test::setIsEndOfTest(bool value)		   { isEndOfTest=value;}
int 	   Test::getAmountOfAnsweredQuestions()    { return amountOfAnsweredQuestions;}
void 	   Test::addAmountOfAnsweredQuestions()    { amountOfAnsweredQuestions++;}

int Test::testResult(){
	double result=(double)(10*collectedPoints)/maxPoints;
	if (result-(int)result>=0.5) {
		return ((int)result)+1;
	}
	else return (int)result;
}

