
#pragma once

class CAddinApp : public CWinApp
{
public:
	void OnCommand(UINT id);

	Visio::IVApplicationPtr GetVisioApp();
	void SetVisioApp(Visio::IVApplicationPtr app);

	virtual BOOL InitInstance();
	virtual int ExitInstance();

private:
	Visio::IVApplicationPtr m_app;
};

extern CAddinApp theApp;
