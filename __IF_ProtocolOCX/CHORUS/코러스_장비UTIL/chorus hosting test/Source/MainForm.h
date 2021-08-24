//---------------------------------------------------------------------------
#ifndef MainFormH
#define MainFormH
//---------------------------------------------------------------------------
#include <Classes.hpp>
#include <Controls.hpp>
#include <StdCtrls.hpp>
#include <Forms.hpp>
#include <ComCtrls.hpp>
#include <ExtCtrls.hpp>
//---------------------------------------------------------------------------
class TMain: public TForm
{
__published:
	TButton* Begin;
	TGroupBox* Info;
	TEdit* InfoSample;
	TLabel* Label1;
	TLabel* Label2;
	TLabel* Label3;
	TLabel* Label4;
	TLabel* Label5;
	TLabel* Label6;
	TLabel* Label7;
	TLabel* Label8;
	TLabel* Label9;
	TLabel* Label10;
	TEdit* MeasureUnit;
	TMemo* Output;
	TCheckBox* Pediatric;
	TGroupBox* Programming;
	TComboBox* ProgramTest;
	TComboBox* Report;
	TButton* Restart;
	TGroupBox* Result;
	TEdit* ResultSample;
	TEdit* ResultTest;
	TGroupBox* SerialParameters;
	TEdit* SerialPort;
	TComboBox* SerialSpeed;
	TStatusBar* Status;
	TEdit* Storable;
	TRadioGroup* Target;
	TEdit* Titration;
	void __fastcall BeginClick(TObject* sender);
	void __fastcall RestartClick(TObject* sender);
	void __fastcall InfoSampleExit(TObject* sender);
	void __fastcall StorableExit(TObject* sender);
	void __fastcall ResultSampleExit(TObject* sender);
	void __fastcall ResultTestExit(TObject* sender);
	void __fastcall TitrationExit(TObject* sender);
	void __fastcall MeasureUnitExit(TObject* sender);
	void __fastcall TargetClick(TObject* sender);
private:
	static void ProgrammingProcessor(std::string sampleCode, bool pedFlag, std::vector<WORD> testList, BYTE& storableRec);
	static void ResultProcessor(std::string measureUnit, char report, std::string sampleCode, std::string test, std::string titration);
	static SampleProgramming SampleProgrammer(std::string sampleCode, BYTE storableRec);
public:
	__fastcall TMain(TComponent* owner) : TForm(owner) { }
};
//---------------------------------------------------------------------------
extern PACKAGE TMain* Main;
//---------------------------------------------------------------------------
#endif
