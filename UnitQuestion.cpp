//---------------------------------------------------------------------------

#pragma hdrstop

#include "UnitQuestion.h"
#include "FormStart.h"
#include "FormCharacteristics.h"
#include "FormTest.h"
#include "UnitTest.h"

#include <vcl.h>
//---------------------------------------------------------------------------
#pragma package(smart_init)

Question::Question(){
	difficulty=1;
	text="";
}

int   	   Question::getDifficulty()						{ return difficulty; 			}
void  	   Question::setDifficulty(int Difficulty)         	{ difficulty=Difficulty;			}
int   	   Question::getImagesAmount()						{ return imagesAmount; 			}
void  	   Question::setImagesAmount(int images)         	{ imagesAmount=images;			}
int        Question::getQuestionNumber()					{ return questionNumber; }
void       Question::setQuestionNumber(int number)			{ questionNumber=number; }
int        Question::getVideoResolutionX()					{ return videoResolutionX; }
void       Question::setVideoResolutionX(int value)			{ videoResolutionX=value; }
int        Question::getVideoResolutionY()					{ return videoResolutionY; }
void       Question::setVideoResolutionY(int value)			{ videoResolutionY=value; }
AnsiString Question::getText()								{ return text; }
void       Question::setText(AnsiString Text)				{ text=Text; }
AnsiString Question::getRightVariant()						{ return rightVariant; }
void       Question::setRightVariant(AnsiString Text)		{ rightVariant=Text; }
AnsiString Question::getPicture(int number)					{ return picture[number];}
void       Question::setPicture(AnsiString way,int number)	{ picture[number]=way;}
AnsiString Question::getUserAnswer()						{ return userAnswer;}
void       Question::setUserAnswer(AnsiString answer)		{ userAnswer=answer;}
AnsiString Question::getQuestionType()						{ return questionType;}
void       Question::setQuestionType(AnsiString type)		{ questionType=type;}
bool       Question::getIsRightAnswer()						{ return isRightAnswer;}
void       Question::setIsRightAnswer(bool isRight)			{ isRightAnswer=isRight;}
bool       Question::getIsAnswered()						{ return isAnswered;}
void       Question::setIsAnswered(bool value)				{ isAnswered=value;}
int 	   Question::getRightVariantsAmount()				{ return rightVariantsAmount;}
void 	   Question::setRightVariantsAmount(int amount)     { rightVariantsAmount=amount;}
AnsiString Question::getUserArrangeAnswer(int index)		{ return userArrangeAnswers[index];}
void       Question::setUserArrangeAnswer(int index, AnsiString value){ userArrangeAnswers[index]=value;}
AnsiString Question::getReference()							{ return reference;}
void       Question::setReference(AnsiString value)			{ reference=value;}
AnsiString Question::getWayToMediaFile()					{ return wayToMediaFile;}
void       Question::setWayToMediaFile(AnsiString value)	{ wayToMediaFile=value;}

AnsiString Question::getVariant(int i){
	if (i >= 0 && i < 10) {  //MAX_VARIANTS_AMOUNT
		return variants[i];
	}
	else return 0;
}

void Question::setVariant(AnsiString Text,int i){
	if (i >= 0 && i < 10) { //MAX_VARIANTS_AMOUNT
		variants[i]=Text;
	}
}
