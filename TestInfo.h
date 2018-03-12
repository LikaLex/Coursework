//---------------------------------------------------------------------------
#pragma once
#ifndef TestInfoH
#define TestInfoH
//---------------------------------------------------------------------------
#endif
const int MAX_THEME_LENGTH=40;
const int MAX_USERNAME_LENGTH=20;

struct TestInfo{
	int amountOfPassings;
	int lastResult;
	int lastTime;
	char user[MAX_USERNAME_LENGTH];
	char theme[MAX_THEME_LENGTH];

	bool isTimeLimited;
	bool areImages;
	int timeIndex;
	int questionsAmount;
	int variantsAmount;
};