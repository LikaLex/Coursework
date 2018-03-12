//---------------------------------------------------------------------------

#ifndef UnitQuestionH
#define UnitQuestionH
//---------------------------------------------------------------------------
#endif
#include <vcl.h>
#define MAXIMAL_VARIANTS_AMOUNT 9
#define MAXIMAL_PICTURES_AMOUNT 4
#define ARRANGE 4
class Question{
	private:
		int difficulty;
		int questionNumber;
		int rightVariantsAmount;
		int imagesAmount;
		int videoResolutionX;
		int videoResolutionY;
		bool isRightAnswer;
		bool isAnswered;
		AnsiString text;
		AnsiString picture[MAXIMAL_PICTURES_AMOUNT];
		AnsiString rightVariant;
		AnsiString variants[MAXIMAL_VARIANTS_AMOUNT];
		AnsiString userArrangeAnswers[ARRANGE];
		AnsiString userAnswer;
		AnsiString questionType;
		AnsiString reference;
		AnsiString wayToMediaFile;

	public:
		Question();

		int        getDifficulty();
		void       setDifficulty(int Difficulty);
		int        getImagesAmount();
		void       setImagesAmount(int images);
		int        getQuestionNumber();
		void       setQuestionNumber(int number);
		int        getVideoResolutionX();
		void       setVideoResolutionX(int value);
		int        getVideoResolutionY();
		void       setVideoResolutionY(int value);
		bool       getIsRightAnswer();
		void 	   setIsRightAnswer(bool isRight);
		bool       getIsAnswered();
		void 	   setIsAnswered(bool value);
		AnsiString getUserAnswer();
		void 	   setUserAnswer(AnsiString answer);
		AnsiString getText();
		void       setText(AnsiString Text);
		AnsiString getRightVariant();
		void       setRightVariant(AnsiString Text);
		AnsiString getVariant(int i);
		void       setVariant(AnsiString Text,int i);
		AnsiString getPicture(int number);
		void       setPicture(AnsiString way,int number);
		AnsiString getQuestionType();
		void       setQuestionType(AnsiString type);
		AnsiString getUserArrangeAnswer(int index);
		void       setUserArrangeAnswer(int index,AnsiString value);
		int        getRightVariantsAmount();
		void       setRightVariantsAmount(int amount);
		AnsiString getReference();
		void       setReference(AnsiString value);
		AnsiString getWayToMediaFile();
		void       setWayToMediaFile(AnsiString value);
};

