//---------------------------------------------------------------------------
#include <vcl.h>
#pragma hdrstop

#include "Chorus.h"

#pragma package(smart_init)

using namespace std;
//---------------------------------------------------------------------------
void Chorus::Enquiry() const
{
	Frame enquiry(Enq);
	enquiry.Write(InstrID);
	Send(enquiry);

	Frame ack = Receive();

	if (ack.Size() != 0 || ack.Command() != Eot)
		throw Exception("Error: bad ack frame.");
}
//---------------------------------------------------------------------------
void Chorus::ProcessJList(vector<string> sampleCodeList, BYTE storableRec, ProgrammingProcessor processor) const
{
	Enquiry();

	for (vector<string>::const_iterator sampleCode = sampleCodeList.begin(); sampleCode != sampleCodeList.end(); sampleCode++)
	{
		Frame request(JListCmd);
		request.Write(storableRec);
		request.WriteString(*sampleCode, 18);
		Send(request);

		Frame response = Receive();

		try
		{
			if (response.Size() < 20 || response.Size() % 2 != 0 || response.Command() != Eot)
				throw Exception("");

			if (ParseString(response, 0, 19).compare(*sampleCode) != 0)
				throw Exception("");

			vector<WORD> testList;
			for (size_t index = 20; index < response.Size(); index += 2)
				testList.push_back(response[index] + (response[index + 1] << 8));

			processor(*sampleCode, response[19], testList, storableRec);
		}
		catch (const Exception&)
		{
			throw Exception("Error: bad response frame.");
		}
	}

	Frame end(JListEnd);
	end.Write(InstrID);
	Send(end);

	Frame ack = Receive();

	if (ack.Size() != 0 || ack.Command() != Eot)
		throw Exception("Error: bad ack frame.");
}
//---------------------------------------------------------------------------
void Chorus::SendResults(vector<SampleResult> resultList) const
{
	Enquiry();

	for (vector<SampleResult>::const_iterator result = resultList.begin(); result != resultList.end(); result++)
	{
		Frame request(ResFrame);
		request.WriteString(result->SampleCode, 18);
		request.WriteString(result->Test, 6);
		request.Write(result->Report);
		request.WriteString(result->Titration, 11);
		request.WriteString(result->MeasureUnit, 9);
		Send(request);

		Frame response = Receive();

		if (response.Size() != 0 || response.Command() != Eot)
			throw Exception("Error: bad response frame.");
	}

	Frame end(ResEnd);
	end.Write(InstrID);
	Send(end);

	Frame ack = Receive();

	if (ack.Size() != 0 || ack.Command() != Eot)
		throw Exception("Error: bad ack frame.");
}
//---------------------------------------------------------------------------
