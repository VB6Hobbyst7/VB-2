//---------------------------------------------------------------------------
#include <sstream>
#include <vcl.h>
#pragma hdrstop

#include "Chorus.h"
#include "Host.h"
#include "MainForm.h"

#pragma package(smart_init)
#pragma resource "*.dfm"

using namespace std;

TMain* Main;
//---------------------------------------------------------------------------
void __fastcall TMain::BeginClick(TObject* sender)
{
	Begin->Enabled = false;

	SerialPort->Enabled = false;
	SerialSpeed->Enabled = false;
	Target->Enabled = false;

	Pediatric->Enabled = false;
	ProgramTest->Enabled = false;

	Storable->Enabled = false;
	InfoSample->Enabled = false;
	ResultSample->Enabled = false;
	ResultTest->Enabled = false;
	Report->Enabled = false;
	Titration->Enabled = false;
	MeasureUnit->Enabled = false;

	Application->ProcessMessages();

	try
	{
		if (Target->ItemIndex == 0)
		{
			Host host(SerialPort->Text.c_str(), SerialSpeed->Text.ToInt());

			Status->SimpleText = "Select C-List -> Host on the Chorus...";
			host.SendJList(SampleProgrammer);

			Status->SimpleText = "Select Arch -> View -> Host on the Chorus...";
			host.ProcessResults(ResultProcessor);
		}
		else
		{
			Chorus chorus(SerialPort->Text.c_str(), SerialSpeed->Text.ToInt());

			Status->SimpleText = "Testing coupling list...";
			vector<string> list;
			list.push_back(InfoSample->Text.c_str());
			chorus.ProcessJList(list, Storable->Text.ToInt(), ProgrammingProcessor);

			Status->SimpleText = "Testing archive...";
			vector<SampleResult> resultList;
			resultList.push_back(SampleResult(
				MeasureUnit->Text.c_str(),
				Report->Text[1],
				ResultSample->Text.c_str(),
				ResultTest->Text.c_str(),
				Titration->Text.c_str()));
			chorus.SendResults(resultList);
		}

		Status->SimpleText = "Test successfully completed.";
	}
	catch (const Exception& ex)
	{
		Status->SimpleText = ex.Message;
	}

	Restart->Enabled = true;
}
//---------------------------------------------------------------------------
void __fastcall TMain::InfoSampleExit(TObject* sender)
{
	if (InfoSample->Text.Length() > 18)
	{
		ShowMessage("Sample code cannot be longer than 18 characters.");
		InfoSample->SelectAll();
		InfoSample->SetFocus();
	}
}
//---------------------------------------------------------------------------
void __fastcall TMain::MeasureUnitExit(TObject* sender)
{
	if (MeasureUnit->Text.Length() > 9)
	{
		ShowMessage("Measure Unit cannot be longer than 9 characters.");
		MeasureUnit->SelectAll();
		MeasureUnit->SetFocus();
	}
}
//---------------------------------------------------------------------------
void TMain::ProgrammingProcessor(string sampleCode, bool pedFlag, vector<WORD> testList, BYTE& storableRec)
{
	ostringstream tests;
	tests << "[";
	for (vector<WORD>::const_iterator test = testList.begin(); test != testList.end(); test++)
		tests << (test == testList.begin() ? "" : ", ") << *test;
	tests << "]";

	Main->Output->Lines->Add(AnsiString("[JList] Sample: ") + sampleCode.c_str() +
		", Pediatric: " + (pedFlag ? "Y" : "N") +
		", Test: " + tests.str().c_str());

	storableRec -= testList.size();
}
//---------------------------------------------------------------------------
void __fastcall TMain::RestartClick(TObject* sender)
{
	Begin->Enabled = true;
	Restart->Enabled = false;

	SerialPort->Enabled = true;
	SerialSpeed->Enabled = true;
	Target->Enabled = true;

	TargetClick(0);

	Output->Lines->Clear();
	Status->SimpleText = "Ready.";
}
//---------------------------------------------------------------------------
void TMain::ResultProcessor(string measureUnit, char report, string sampleCode, string test, string titration)
{
	Main->Output->Lines->Add(AnsiString("[Result] Sample: ") + sampleCode.c_str() +
		", Test: " + test.c_str() +
		", Result: " + titration.c_str() + " " + measureUnit.c_str() + " (" + report + ")");
}
//---------------------------------------------------------------------------
void __fastcall TMain::ResultSampleExit(TObject* sender)
{
	if (ResultSample->Text.Length() > 18)
	{
		ShowMessage("Sample code cannot be longer than 18 characters.");
		ResultSample->SelectAll();
		ResultSample->SetFocus();
	}
}
//---------------------------------------------------------------------------
void __fastcall TMain::ResultTestExit(TObject* sender)
{
	if (ResultTest->Text.Length() > 6)
	{
		ShowMessage("Test cannot be longer than 6 characters.");
		ResultTest->SelectAll();
		ResultTest->SetFocus();
	}
}
//---------------------------------------------------------------------------
SampleProgramming TMain::SampleProgrammer(string sampleCode, BYTE storableRec)
{
	Main->Output->Lines->Add(AnsiString("[JList] Storable: ") + storableRec +
		", Sample: " + sampleCode.c_str());

	SampleProgramming programming(Main->Pediatric->Checked);
	programming.TestList.push_back(Main->ProgramTest->ItemIndex + 1);
	return programming;
}
//---------------------------------------------------------------------------
void __fastcall TMain::StorableExit(TObject* sender)
{
	try
	{
		int storable = Storable->Text.ToInt();
		if (storable < 1 || storable > 255)
			throw true;
	}
	catch (...)
	{
		ShowMessage("Storable records must be a number between 1 and 255.");
		Storable->SelectAll();
		Storable->SetFocus();
	}
}
//---------------------------------------------------------------------------
void __fastcall TMain::TargetClick(TObject* sender)
{
	Pediatric->Enabled = Target->ItemIndex == 0;
	ProgramTest->Enabled = Target->ItemIndex == 0;

	Storable->Enabled = Target->ItemIndex == 1;
	InfoSample->Enabled = Target->ItemIndex == 1;
	ResultSample->Enabled = Target->ItemIndex == 1;
	ResultTest->Enabled = Target->ItemIndex == 1;
	Report->Enabled = Target->ItemIndex == 1;
	Titration->Enabled = Target->ItemIndex == 1;
	MeasureUnit->Enabled = Target->ItemIndex == 1;
}
//---------------------------------------------------------------------------
void __fastcall TMain::TitrationExit(TObject* sender)
{
	if (Titration->Text.Length() > 11)
	{
		ShowMessage("Titration cannot be longer than 11 characters.");
		Titration->SelectAll();
		Titration->SetFocus();
	}
}
//---------------------------------------------------------------------------
