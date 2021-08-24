//---------------------------------------------------------------------------
#include <vcl.h>
#pragma hdrstop

#include "Handler.h"

#pragma package(smart_init)

using namespace std;
//---------------------------------------------------------------------------
Handler::Handler(const char* port, int speed, int timeout)
{
	if ((handle = CreateFile(port, GENERIC_READ | GENERIC_WRITE, 0, 0, OPEN_EXISTING, 0, 0)) == INVALID_HANDLE_VALUE)
		throw Exception("Error: cannot open serial port.");

	DCB dcb;
	dcb.DCBlength = sizeof(dcb);
	if (!GetCommState(handle, &dcb))
		throw Exception("Unexpected error: GetCommState().");

	dcb.BaudRate = speed;
	dcb.ByteSize = 8;
	dcb.Parity = NOPARITY;
	dcb.StopBits = ONESTOPBIT;
	dcb.fBinary = true;
	dcb.fRtsControl = RTS_CONTROL_DISABLE;
	dcb.fDtrControl = DTR_CONTROL_DISABLE;
	if (!SetCommState(handle, &dcb))
		throw Exception("Unexpected error: SetCommState().");

	COMMTIMEOUTS ct;
	ct.ReadIntervalTimeout = 0;
	ct.ReadTotalTimeoutMultiplier = 0;
	ct.ReadTotalTimeoutConstant = timeout;
	ct.WriteTotalTimeoutMultiplier = 0;
	ct.WriteTotalTimeoutConstant = 0;
	if (!SetCommTimeouts(handle, &ct))
		throw Exception("Unexpected error: SetCommTimeouts().");
}
//---------------------------------------------------------------------------
Handler::~Handler()
{
	if (handle != INVALID_HANDLE_VALUE)
		CloseHandle(handle);
}
//---------------------------------------------------------------------------
string Handler::ParseString(const Frame& frame, int position, int length) const
{
	string text = "";
	int index = position;
	while ((index < position + length) && frame[index])
		text += frame[index++];

	if (index == position + length)
		throw Exception("Bad frame.");

	return text;
}
//---------------------------------------------------------------------------
Frame Handler::Receive() const
{
	BYTE data[258];
	DWORD read;

	do
	{
		if (!ReadFile(handle, data, 1, &read, 0) || read != 1)
			throw Exception("Error receiving data.");
	}
	while (data[0] != Stx);

	if (!ReadFile(handle, data + 1, 1, &read, 0) || read != 1)
		throw Exception("Error receiving data.");

	if (data[1] < 1)
		throw Exception("Error: frame too short.");

	if (!ReadFile(handle, data + 2, data[1] + 1, &read, 0) || read != (DWORD)(data[1] + 1))
		throw Exception("Error receiving data.");

	return Frame(data);
}
//---------------------------------------------------------------------------
void Handler::Send(Frame& frame) const
{
	DWORD written;

	if (!WriteFile(handle, frame.Stream(), frame.Size() + 4, &written, 0) || written != frame.Size() + 4)
		throw Exception("Error sending data.");
}
//---------------------------------------------------------------------------
