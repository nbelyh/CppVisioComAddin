#include "stdafx.h"
#include "AddinCommands.h"

/*
	Define the add-in command tags (ids) for convenience
*/

const LPCWSTR COMMAND_1 = L"$VCXPROJECTNAME$.button1";
const LPCWSTR COMMAND_2 = L"$VCXPROJECTNAME$.button2";


/*
	Constructor
*/

CAddinCommands::CAddinCommands(Visio::IVApplicationPtr app)
	: m_visio(app)
{
}

/*
	The function is called when user clicks corresponding button.
	The main processing should happen here.
*/

void CAddinCommands::OnCommand(LPCWSTR cmd) const
{
	MessageBoxW(NULL, L"A button was clicked!", cmd, MB_OK | MB_ICONINFORMATION);
}

/*
	Return image aassociated with the command
*/

UINT CAddinCommands::GetCommandIconId(LPCWSTR cmd) const
{
	if (!StrCmpI(cmd, COMMAND_1)) return IDI_ICON1;
	if (!StrCmpI(cmd, COMMAND_2)) return IDI_ICON2;
	return 0;
}

/*
	Return command name
*/

bstr_t CAddinCommands::GetCommandLabel(LPCWSTR cmd) const
{
	if (!StrCmpI(cmd, COMMAND_1)) return "button1";
	if (!StrCmpI(cmd, COMMAND_2)) return "button2";
	return L"";
}

/* 
	Configure toolbar stuff for old Visio
*/

void CAddinCommands::BuildToolbar(IToolbarBuilder* builder) const
{
	builder->AddToolbar(L"$VCXPROJECTNAME$");
	builder->AddButton(COMMAND_1);
	builder->AddButton(COMMAND_2);
}
