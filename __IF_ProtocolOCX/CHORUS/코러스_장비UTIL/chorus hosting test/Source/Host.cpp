//---------------------------------------------------------------------------
#include <vcl.h>
#pragma hdrstop

#include "Host.h"

#pragma package(smart_init)

using namespace std;
//---------------------------------------------------------------------------
void Host::Enquiry() const
{
	Frame enquiry = Receive();

	if (enquiry.Size() != 1 || enquiry.Command() != Enq || enquiry[0] != InstrID)
		throw Exception("Error: bad enquiry frame.");

	Send(Frame(Eot));
}
//---------------------------------------------------------------------------
bool Host::ProcessResult(ResultProcessor processor) const
{
	Frame request = Receive();
	switch (request.Command())
	{
		case ResFrame:
			try
			{
				if (request.Size() != 49)
					throw Exception("Error: bad result frame.");

				processor(ParseString(request, 39, 49),
					request[26],
					ParseString(request, 0, 19),
					ParseString(request, 19, 26),
					ParseString(request, 27, 39));

				Send(Frame(Eot));

				return true;
			}
			catch (const Exception&)
			{
				throw Exception("Error: bad JList frame.");
			}
		case ResEnd:
			if (request.Size() != 1 || request[0] != InstrID)
				throw Exception("Error: bad result frame.");

			Send(Frame(Eot));

			return false;
		default:
			throw Exception("Error: bad result frame.");
	}
}
//---------------------------------------------------------------------------
void Host::ProcessResults(ResultProcessor processor) const
{
	Enquiry();

	while (ProcessResult(processor));
}
//---------------------------------------------------------------------------
bool Host::ProgramSample(SampleProgrammer programmer) const
{
	Frame request = Receive();
	switch (request.Command())
	{
		case JListCmd:
			try
			{
				if (request.Size() != 20 || request[0] < 1)
					throw Exception("");

				string sampleCode = ParseString(request, 1, 19);
				SampleProgramming programming = programmer(sampleCode, request[0]);

				Frame response(Eot);
				response.WriteString(sampleCode, 18);
				response.Write(programming.PedFlag);
				for (vector<WORD>::const_iterator test = programming.TestList.begin(); test != programming.TestList.end(); test++)
					response.WriteWord(*test);
				Send(response);

				return true;
			}
			catch (const Exception&)
			{
				throw Exception("Error: bad JList frame.");
			}
		case JListEnd:
			if (request.Size() != 1 || request[0] != InstrID)
				throw Exception("Error: bad JList frame.");

			Send(Frame(Eot));

			return false;
		default:
			throw Exception("Error: bad JList frame.");
	}
}
//---------------------------------------------------------------------------
void Host::SendJList(SampleProgrammer programmer) const
{
	Enquiry();

	while (ProgramSample(programmer));
}
//---------------------------------------------------------------------------
