//---------------------------------------------------------------------------
#ifndef ChorusH
#define ChorusH
//---------------------------------------------------------------------------
#include "Handler.h"
//---------------------------------------------------------------------------
class SampleResult
{
	public:
		std::string MeasureUnit;
		char Report;
		std::string SampleCode;
		std::string Test;
		std::string Titration;

		SampleResult(std::string measureUnit, char report, std::string sampleCode, std::string test, std::string titration) :
			MeasureUnit(measureUnit), Report(report), SampleCode(sampleCode), Test(test), Titration(titration) { }
};
//---------------------------------------------------------------------------
typedef void (*ProgrammingProcessor)(std::string sampleCode, bool pedFlag, std::vector<WORD> testList, BYTE& storableRec);
//---------------------------------------------------------------------------
class Chorus : public Handler
{
	void Enquiry() const;

	public:
		Chorus(const char* port, int speed) : Handler(port, speed, 2000) { }

		void ProcessJList(std::vector<std::string> sampleCodeList, BYTE storableRec, ProgrammingProcessor processor) const;
		void SendResults(std::vector<SampleResult> resultList) const;
};
//---------------------------------------------------------------------------
#endif
