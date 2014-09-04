
#pragma once

struct IToolbarBuilder
{
	virtual void AddToolbar(LPCWSTR name) = 0;
	virtual void AddButton(LPCWSTR cmd) = 0;
};

class CAddinCommands
{
public:
	CAddinCommands(Visio::IVApplicationPtr app);

	void OnCommand(LPCWSTR cmd) const;
	UINT GetCommandIconId(LPCWSTR cmd) const;
	bstr_t GetCommandLabel(LPCWSTR cmd) const;
	
	void BuildToolbar(IToolbarBuilder* builder) const;

private:
	Visio::IVApplicationPtr m_visio;
};
