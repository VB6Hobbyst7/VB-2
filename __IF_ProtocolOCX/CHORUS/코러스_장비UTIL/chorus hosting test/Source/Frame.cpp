//---------------------------------------------------------------------------
#include <vcl.h>
#pragma hdrstop

#include "Frame.h"

#pragma package(smart_init)

using namespace std;
//---------------------------------------------------------------------------
Frame::Frame(const BYTE* input) : command(input[2])
{
	BYTE checksum = input[1] ^ input[2];

	for (int index = 0; index < input[1] - 1; index++)
	{
		checksum ^= input[index + 3];
		data.push_back(input[index + 3]);
	}

	if (checksum != input[input[1] + 2])
		throw Exception("Error: wrong checksum.");
}
//---------------------------------------------------------------------------
const BYTE* Frame::Stream()
{
	stream[0] = Stx;
	stream[1] = Size() + 1;
	stream[2] = command;

	BYTE checksum = stream[1] ^ stream[2];

	for (size_t index = 0; index < Size(); index++)
	{
		checksum ^= data[index];
		stream[index + 3] = data[index];
	}

	stream[Size() + 3] = checksum;

	return stream;
}
//---------------------------------------------------------------------------
void Frame::WriteString(string value, size_t maxLength)
{
	for (size_t index = 0; index < maxLength; index++)
		Write(value.size() > index ? value[index] : 0);
	Write(0);
}
//---------------------------------------------------------------------------
void Frame::WriteWord(WORD value)
{
	Write(value);
	Write(value >> 8);
}
//---------------------------------------------------------------------------
