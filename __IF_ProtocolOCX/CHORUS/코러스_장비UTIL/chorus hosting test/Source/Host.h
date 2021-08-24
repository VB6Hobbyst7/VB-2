//---------------------------------------------------------------------------
#ifndef HostH
#define HostH
//---------------------------------------------------------------------------
#include "Handler.h"
//---------------------------------------------------------------------------
class SampleProgramming
{
	public:
		bool PedFlag;
		std::vector<WORD> TestList;

		SampleProgramming(bool pedFlag) : PedFlag(pedFlag) { }
};
//---------------------------------------------------------------------------
typedef void (*ResultProcessor)(std::string measureUnit, char report, std::string sampleCode, std::string test, std::string titration);
typedef SampleProgramming (*SampleProgrammer)(std::string sampleCode, BYTE storableRec);
//---------------------------------------------------------------------------
class Host : public Handler
{
	void Enquiry() const;
	bool ProcessResult(ResultProcessor processor) const;
	bool ProgramSample(SampleProgrammer programmer) const;

	public:
		Host(const char* port, int speed) : Handler(port, speed, 20000) { }

		void ProcessResults(ResultProcessor processor) const;
		void SendJList(SampleProgrammer programmer) const;
};
//---------------------------------------------------------------------------
#endif
