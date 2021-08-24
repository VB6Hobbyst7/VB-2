//---------------------------------------------------------------------------
#ifndef HandlerH
#define HandlerH
//---------------------------------------------------------------------------
#include <windows.h>

#include "Frame.h"
//---------------------------------------------------------------------------
class Handler
{
	HANDLE handle;

	protected:
		Handler(const char* port, int speed, int timeout);

		std::string ParseString(const Frame& frame, int position, int length) const;
		Frame Receive() const;
		void Send(Frame& frame) const;

	public:
		~Handler();
};
//---------------------------------------------------------------------------
#endif
