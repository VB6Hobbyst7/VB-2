//---------------------------------------------------------------------------
#ifndef FrameH
#define FrameH
//---------------------------------------------------------------------------
#include <vector>
#include <windows.h>

enum
{
	Enq = 0x05,
	Eot = 0x04,
	InstrID = 0x43,
	JListCmd = 0xD2,
	JListEnd = 0xD3,
	ResEnd = 0xD8,
	ResFrame = 0xD7,
	Stx = 0x02
};
//---------------------------------------------------------------------------
class Frame
{
	BYTE command;
	std::vector<BYTE> data;
	BYTE stream[258];

	public:
		Frame(BYTE command) : command(command) { }
		Frame(const BYTE* input);

		BYTE operator [](int index) const { return data[index]; }

		BYTE Command() const { return command; }
		size_t Size() const { return data.size(); }
		const BYTE* Stream();
		void Write(BYTE value) { data.push_back(value); }
		void WriteString(std::string value, size_t maxLength);
		void WriteWord(WORD value);
};
//---------------------------------------------------------------------------
#endif
