//---------------------------------------------------------------------------
#ifndef ExceptionH
#define ExceptionH
//---------------------------------------------------------------------------
#include <string>
//---------------------------------------------------------------------------
class Exception : public std::exception
{
	std::string message;

	public:
		Exception(std::string message) : message(message) { }

		virtual const char* what() const { return message.c_str(); }
};
//---------------------------------------------------------------------------
#endif
