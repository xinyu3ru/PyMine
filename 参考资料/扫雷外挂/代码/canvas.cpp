
#include "stdafx.h"
#include "resource.h"
//#include <afxwin.h>
#include <assert.h>
#include <stdio.h>
#include <commctrl.h>
#include <time.h>
#include <SHELLAPI.H>

#pragma comment(lib, "COMCTL32.LIB")

// DEFINE.
const int NUM_BITMAPS = 16;

const int BOMB_UNCLK  = 0;
const int BOMB_1      = 1;
const int BOMB_2      = 2;
const int BOMB_3      = 3;
const int BOMB_4      = 4;
const int BOMB_5      = 5;
const int BOMB_6      = 6;
const int BOMB_7      = 7;
const int BOMB_8      = 8;
const int BOMB_BLANK  = 9;
const int BOMB_BOMB   = 10;
const int BOMB_UNKOWN = 11;
const int BOMB_NOBOMB = 12;
const int BOMB_FLAG   = 13;
const int BOMB_CLKED  = 14;
const int BOMB_DICEYD = 15;

//{{FREE CODE
/////////////////////////////////////////////////////////////////////////////////////////
HINSTANCE hInst;
HDC		  hdc, mdc;
HWND	  hWnd, g_WindowHWnd;
HWND      hStatusWnd, hToolsWnd;
HBITMAP   map[NUM_BITMAPS];
DWORD     tBegin=0, tNow=0;
HMENU     hMenu ;

LPCSTR    strAbout = "All Rights Reserved By COX @ SWUST. (C)2007";

void     BombInit();
void     BombMain();
void     UpdateState();
void     GetBombPos();
void     AIClick();
void     ReDo();
void     RandClick();
void     SpeedRun();
void     AIRun();
int      GetInfo(int row, int line );
COLORREF GetColor( int newX, int newY );
void     MouseDown(LPARAM lParam);
// void     assert( void* param );

typedef enum tagGameState { GAME_BEGIN, GAME_PAUSE, GAME_OVER, GAME_AI, GAME_AIRUN };
tagGameState GameState = GAME_BEGIN;

// 游戏左上角坐标
const int GAME_LEFT   = 12;
const int GAME_TOP    = 55;
const int GAME_RIGHT  = 492;
const int GAME_BOTTOM = 311;

const int BOMB_WIDTH  = 16;
const int BOMB_HEIGHT = 16;
const int BOMB_ROW    = 16;
const int BOMB_LINE   = 30;
int BombState[BOMB_ROW][BOMB_LINE];
 
/////////////////////////////////////////////////////////////////////////////////////////
ATOM				MyRegisterClass( HINSTANCE hInstance );
BOOL				InitInstance( HINSTANCE, int );
LRESULT CALLBACK	WndProc( HWND, UINT, WPARAM, LPARAM );
int     CALLBACK    DlgAboutProc( HWND, UINT, WPARAM, LPARAM );

//////////////////////////////////////////////////////////////////////////////
ATOM MyRegisterClass( HINSTANCE hInstance )
{
	WNDCLASSEX wcex;

	wcex.cbSize			= sizeof( WNDCLASSEX ); 
	wcex.style			= CS_HREDRAW | CS_VREDRAW;
	wcex.lpfnWndProc	= ( WNDPROC )WndProc;
	wcex.cbClsExtra		= 0;
	wcex.cbWndExtra		= 0;
	wcex.hInstance		= hInstance;
	wcex.hIcon			= ( HICON )LoadIcon( hInst, MAKEINTRESOURCE(IDI_ICON) );
	wcex.hCursor		= ( HCURSOR )LoadCursor( NULL, IDC_ARROW );
	wcex.hbrBackground	= ( HBRUSH )GetStockObject( WHITE_BRUSH );
	wcex.lpszMenuName	= NULL;
	wcex.lpszClassName	= "canvas";
	wcex.hIconSm		= ( HICON )LoadIcon( hInst, MAKEINTRESOURCE(IDI_ICON_SMALL) );

	return RegisterClassEx( &wcex );
}

////////////////////////////////////////////////////////////////////////////////////
//the windows main function
int APIENTRY WinMain( HINSTANCE hInstance,
                      HINSTANCE hPrevInstance,
                      LPSTR     lpCmdLine,
                      int       nCmdShow )
{
	MSG msg;

	MyRegisterClass( hInstance );

	//register windows class and instance
	if( !InitInstance( hInstance, nCmdShow ) ) 
	{
		return FALSE;
	}

	//deal with the windows messages
	PeekMessage(&msg, NULL, 0, 0, PM_NOREMOVE);
    while( msg.message != WM_QUIT )
    {
        if( PeekMessage( &msg, NULL, 0, 0, PM_REMOVE ) )
        {
            TranslateMessage( &msg );
            DispatchMessage( &msg );
        }
		else
		{
			if ( GameState == GAME_BEGIN )
			{
				BombMain();
			}
			else if ( GameState == GAME_OVER )
			{
				//BombInit();
				MessageBox( 0, "搞定啦~", "扫雷外挂", 0 );
				GameState = GAME_BEGIN;
			}
			else if ( GameState == GAME_AI )
			{
				GameState = GAME_BEGIN;
				for ( int y=0; y<BOMB_ROW; y++ )
				{
					for ( int x=0; x<BOMB_LINE; x++ )
					{
						if ( BombState[y][x] == BOMB_NOBOMB )
						{
							GameState = GAME_AI;
						}
					}
				}
				AIClick();
				BombMain();
			}
			else if ( GameState == GAME_AIRUN )
			{
				tNow = GetTickCount();
				GameState = GAME_OVER;
				bool useAI = false;
				for ( int y=0; y<BOMB_ROW; y++ )
				{
					for ( int x=0; x<BOMB_LINE; x++ )
					{
						if ( BombState[y][x] == BOMB_NOBOMB )
						{
							useAI = true;
						}
						if ( BombState[y][x] == BOMB_BOMB || BombState[y][x] == BOMB_UNCLK )
						{
							GameState = GAME_AIRUN;
						}
					}
				}
				if ( useAI )
				{
					AIClick();
					BombMain();
				}
				else
				{
					RandClick();
					BombMain();
				}
				int total=0;
				for ( y=0; y<BOMB_ROW; y++ )
				{
					for ( int x=0; x<BOMB_LINE; x++ )
					{
						if ( BombState[y][x]==BOMB_BOMB )
						{
							total++;
						}
					}
				}
				if ( total == 99 )
				{
					//ReDo();
					AIClick();
					//GameState = GAME_OVER;
				}
			}
			else
			{
				// GAME_PAUSE
			}
		}
    }

	return msg.wParam;
}

//////////////////////////////////////////////////////////////////////////////////////////
//windows instance
BOOL InitInstance( HINSTANCE hInstance, int nCmdShow )
{
	hInst = hInstance;

	hWnd = CreateWindow( "canvas", "扫雷外挂 [初始化完毕]", 
		WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU | WS_MINIMIZEBOX,
		0, 0, 505, 380, NULL, NULL, hInstance, NULL );
    assert( hWnd );
	
	// Load Menu.
	hMenu = LoadMenu( NULL, MAKEINTRESOURCE( IDR_MENU ) );
	assert( hMenu );
	SetMenu( hWnd, hMenu );
	
	// Load Toolbar.
	hToolsWnd = CreateWindowEx( WS_EX_TOOLWINDOW, 
		TOOLBARCLASSNAME, 
		NULL, 
		WS_CHILD | WS_VISIBLE,
		0, 0, 0, 0,
		hWnd,
		NULL,
		hInstance,
		NULL );
	assert( hToolsWnd );
    HIMAGELIST hImgList = ImageList_Create( 16, 16, ILC_COLOR4, 0, 7 );
	assert( hImgList );
    ImageList_Add( hImgList, LoadBitmap( hInstance, MAKEINTRESOURCE(IDB_BITMAP) ), NULL );   
    SendMessage( hToolsWnd, TB_SETIMAGELIST, 0, (LPARAM)hImgList );   
    TBBUTTON pToolbar[7] = {
		{ 0, ID_TOOLBAR_RUN,    TBSTATE_ENABLED, TBSTYLE_BUTTON, 0, 0 }, 
		{ 1, ID_TOOLBAR_PAUSE,  TBSTATE_ENABLED, TBSTYLE_BUTTON, 0, 0 },
		{ 2, ID_TOOLBAR_STOP,   TBSTATE_ENABLED, TBSTYLE_BUTTON, 0, 0 },
		{ 3, ID_TOOLBAR_RAND,   TBSTATE_ENABLED, TBSTYLE_BUTTON, 0, 0 },
		{ 4, ID_TOOLBAR_SPRUN,  TBSTATE_ENABLED, TBSTYLE_BUTTON, 0, 0 },
		{ 5, ID_TOOLBAR_AI,     TBSTATE_ENABLED, TBSTYLE_BUTTON, 0, 0 },
		{ 6, ID_TOOLBAR_GETTAG, TBSTATE_ENABLED, TBSTYLE_BUTTON, 0, 0 }
	};
    assert( pToolbar );
    SendMessage( hToolsWnd, TB_BUTTONSTRUCTSIZE, sizeof(TBBUTTON), 0 );     
    SendMessage( hToolsWnd, TB_ADDBUTTONS, (UINT)7, (LPARAM)pToolbar );     
    ShowWindow( hToolsWnd, SW_SHOWMAXIMIZED );
	
	// Load Statubar.
	hStatusWnd = CreateStatusWindow( WS_CHILD|WS_VISIBLE|WS_BORDER, NULL, hWnd, IDS_STATUS );
	assert( hStatusWnd );
	int StatuParts[3] = { 250, 350, -1 };
    SendMessage( hStatusWnd, SB_SETPARTS, 3, (LPARAM)StatuParts );
    // SendMessage( hStatusWnd, SB_SETTEXT, 0, (LPARAM)TEXT("状态栏信息") );
	SendMessage( hStatusWnd, SB_SETTEXT, 2, (LPARAM)TEXT("E-mail: cox0902@163.com") );

	ShowWindow( hWnd, SW_SHOWDEFAULT );
	UpdateWindow( hWnd );

	hdc = GetDC( hWnd );
	assert( hdc );

	HWND myWND = FindWindowEx( hWnd, 0, "SYSMENU", NULL );
	HDC  myDC  = GetDC( myWND );
	DrawIcon( myDC, 0, 0, LoadIcon(hInstance, MAKEINTRESOURCE(IDI_ICON_SMALL)) );
	// Make Map.
	HDC     hdcBombs  = CreateCompatibleDC( hdc );
	HBITMAP hbmpBombs = LoadBitmap( hInstance, MAKEINTRESOURCE(IDB_BOMBS) );
	HDC     hdcMaps   = CreateCompatibleDC( hdc );

	SelectObject( hdcBombs, hbmpBombs );
	for ( int i=0; i<NUM_BITMAPS; i++ )
	{
		map[i] = CreateCompatibleBitmap( hdc, BOMB_WIDTH, BOMB_HEIGHT );
		SelectObject( hdcMaps, map[i] );
		BitBlt( hdcMaps, 0, 0, BOMB_WIDTH, BOMB_HEIGHT, hdcBombs, 0, BOMB_HEIGHT*i, SRCCOPY );
		assert( map[i] );
	}
	DeleteDC( hdcBombs );
	DeleteDC( hdcMaps );
	DeleteObject( hbmpBombs );

	BombInit();

	return TRUE;
}

//////////////////////////////////////////////////////////////////////
LRESULT CALLBACK WndProc( HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam )
{
	switch( message )
	{
	case WM_COMMAND:
		switch ( wParam )
		{
		case ID_TOOLBAR_RUN:
			SetWindowText( hWnd, "扫雷外挂 [手动扫雷中]" );
			AIClick();
			break;
		case ID_GAME_CLEAR:
		case ID_TOOLBAR_STOP:
			ReDo();
			break;
		case ID_TOOLBAR_RAND:
			SetWindowText( hWnd, "扫雷外挂 [手动扫雷中]" );
			RandClick();
			break;
		case ID_TOOLBAR_SPRUN:
			SetWindowText( hWnd, "扫雷外挂 [手动扫雷中]" );
			SpeedRun();
			break;
		case ID_GAME_BEGIN:
		case ID_TOOLBAR_AI:
			SetWindowText( hWnd, "扫雷外挂 [自动扫雷中]" );
			if ( GameState != GAME_AIRUN )
			{
				tBegin = GetTickCount();
			}
			AIRun();
			break;
		case ID_GAME_PAUSE:
		case ID_TOOLBAR_PAUSE:
			SetWindowText( hWnd, "扫雷外挂 [暂停中]" );
			GameState=GAME_BEGIN;
			break;
		case ID_ABOUT_RIGHTS:
			ShellAbout( hWnd, "扫雷外挂 v1.1.1", strAbout, LoadIcon( hInst, MAKEINTRESOURCE(IDI_ICON) ));
			break;
		case ID_GAME_EXIT:
			SetWindowPos( g_WindowHWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE );
			DeleteDC( mdc );
			ReleaseDC( hWnd, hdc );
			PostQuitMessage( 0 );
			break;
		case ID_TOOLBAR_GETTAG:
		case ID_SETUP_GET:
			BombInit();
			break;
		default:
			// MessageBox( 0, "Unknown Msg Handle!!", "DEBUG", 0 );
			break;
		}
		break;
	case WM_LBUTTONDOWN:
		MouseDown(lParam);
		break;
	case WM_DESTROY:					///////////////////
		SetWindowPos( g_WindowHWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE );
		DeleteDC( mdc );
		ReleaseDC( hWnd, hdc );
		PostQuitMessage( 0 );
		break;
	default:							//////////////////////
		return DefWindowProc( hWnd, message, wParam, lParam );
   }
   return 0;
}

// My CODE here.
void MouseDown(LPARAM lParam)
{
	RECT rt;
	rt.left   = GAME_LEFT;
	rt.top    = GAME_TOP-BOMB_HEIGHT;
	rt.right  = GAME_RIGHT;//-BOMB_WIDTH;
	rt.bottom = GAME_BOTTOM-BOMB_HEIGHT;
	int x = (lParam&0x0000FFFF)>>0;
	int y = (lParam&0xFFFF0000)>>16;
	if (rt.left <= x && rt.right >= x &&
		rt.top  <= y && rt.bottom >= y)
	{
		// MessageBox( 0, "IN RECT!!", "DEBUG", 0 );
		int kX = x;
		int kY = y+BOMB_HEIGHT;
		// if (GetInfo(kY%16, kX%16) != BOMB_BOMB)
		{
			LPARAM lPar=(kY<<16)|kX;
			PostMessage( g_WindowHWnd, WM_LBUTTONDOWN, 0, lPar );
			PostMessage( g_WindowHWnd, WM_LBUTTONUP, 0, lPar );
		}
	}
}

void BombInit()
{
	WINDOWPLACEMENT winWndPl;	
    g_WindowHWnd = FindWindow( NULL, "扫雷" );
	if ( g_WindowHWnd == NULL )
	{
		if ( ERROR_FILE_NOT_FOUND == WinExec( "winmine.exe", SW_SHOW ) )
		{
			char winminePath[1024] = {'\0'};
			GetSystemDirectory( winminePath, 1024 );
			// MessageBox(NULL, winminePath, "DEBUG", 0);
			strcat( winminePath, "\\winmine.exe" );
			// MessageBox(NULL, winminePath, "DEBUG", 0);
			if ( ERROR_FILE_NOT_FOUND == WinExec( winminePath, SW_SHOW ) )
			{
				MessageBox( NULL, "未能找到扫雷程序!", "扫雷外挂", 0 );
				PostQuitMessage(0);
			}
		}
		g_WindowHWnd = FindWindow( NULL, "扫雷" );
	}
	
    // 窗口总在前
    SetWindowPos( g_WindowHWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE );
    
	GetWindowPlacement( g_WindowHWnd, &winWndPl );
	winWndPl.showCmd = SW_SHOWMINIMIZED;
	SetWindowPlacement( g_WindowHWnd, &winWndPl );
	winWndPl.showCmd = SW_SHOWNORMAL;
	SetWindowPlacement( g_WindowHWnd, &winWndPl );

	// PostMessage( g_WindowHWnd, WM_KEYDOWN, VK_F2, 0 );
}

void    BombMain()
{    
    UpdateState();
    // gameAI();
}

int GetInfo( int row, int line )
{
    COLORREF mycolor[3];
    mycolor[0] = GetColor( line * 16 + GAME_LEFT + 8, row * 16 + GAME_TOP + 8 );
	mycolor[1] = GetColor( line * 16 + GAME_LEFT + 1, row * 16 + GAME_TOP + 8 );
    mycolor[2] = GetColor( line * 16 + GAME_LEFT + 8, row * 16 + GAME_TOP + 4 );

    if ( mycolor[0] == 12632256 )
	{
        if( mycolor[1] == 16777215 )
		{
            return BOMB_UNCLK;     // 未打开
		}
        else
		{
			if ( mycolor[2] == 0 )
			{
				return BOMB_7;
			}
			else
			{
				return BOMB_BLANK;     // 空白
			}
        }
	}
    else if ( mycolor[0] == 16711680 )  // 1个
	{
        return BOMB_1;
	}
    else if ( mycolor[0] == 32768 )     // 2个
	{
        return BOMB_2;
	}
    else if ( mycolor[0] == 255 )       // 3个
	{
		return BOMB_3;
	}
    else if ( mycolor[0] == 8388608 )   // 4个
	{
		return BOMB_4;
	}
    else if ( mycolor[0] == 128 )       // 5个
	{
		return BOMB_5;
	}
    else if ( mycolor[0] == 8421376 )   // 6个
	{
		return BOMB_6;
	}
    else if ( mycolor[0] == 0 )     // 地雷
	{
		if ( mycolor[2] == 255 )
		{
			return BOMB_FLAG;
		}
		PostMessage( g_WindowHWnd, WM_KEYDOWN, VK_F2, 0 );
        return BOMB_BOMB;
	}
	else if( mycolor[0] == 8421504 )
    {
		return BOMB_8;
	}
    else
	{
		//char chTest[1024];
		//sprintf( chTest, "Unknown color!! %d", mycolor );
        //MessageBox( 0, chTest, "DEBUG", 0 );
        return BOMB_UNKOWN;
    }
}

COLORREF GetColor( int newX, int newY )
{
    HDC windowDC;
    RECT winRECT;      // 游戏窗口的RECT
    POINT po;          // 位置
	COLORREF retColor; 
    
    // 获取游戏场景
    windowDC = GetDC(g_WindowHWnd);
	assert( windowDC );
    // 获得游戏窗口的RECT
    GetClientRect( g_WindowHWnd, &winRECT );
	po.x = winRECT.left + newX;
    po.y = winRECT.top + newY;
    retColor = GetPixel( windowDC, po.x, po.y );
    // 释放场景
    ReleaseDC( g_WindowHWnd, windowDC );

	return retColor;
}

void     UpdateState()
{

	for ( int y=0; y<BOMB_ROW; y++ )
	{
		for ( int x=0; x<BOMB_LINE; x++ )
		{
			BombState[y][x]=GetInfo( y, x );
		}
	}
	GetBombPos();
	int total=0;
	char strStatu[64];
	for ( y=0; y<BOMB_ROW; y++ )
	{
		for ( int x=0; x<BOMB_LINE; x++ )
		{
			if ( BombState[y][x] == BOMB_BOMB )
			{
				total++;
			}
		}
	}
	sprintf( strStatu, "当前有%d/(99)个雷已知.", total );
	SendMessage( hStatusWnd, SB_SETTEXT, 0, (LPARAM)strStatu );
	sprintf( strStatu, "耗时: %ds", tNow/1000-tBegin/1000 );
	SendMessage( hStatusWnd, SB_SETTEXT, 1, (LPARAM)strStatu );
	
	mdc = CreateCompatibleDC(hdc);
	assert( mdc );
	HDC     bufdc = CreateCompatibleDC(hdc);
	assert( bufdc );
	HBITMAP fullmap = CreateCompatibleBitmap( hdc, BOMB_LINE*BOMB_WIDTH, BOMB_ROW*BOMB_HEIGHT );
	assert( fullmap );

	SelectObject( mdc, fullmap );

	// mapIndex
	for ( y=0; y<BOMB_ROW; y++ )
	{
		for ( int x=0; x<BOMB_LINE; x++ )
		{
			SelectObject( bufdc, map[BombState[y][x]] );
			BitBlt( mdc, x*BOMB_WIDTH, y*BOMB_HEIGHT, BOMB_WIDTH, BOMB_HEIGHT, bufdc, 0, 0, SRCCOPY );
		}
	}

	SelectObject( mdc,fullmap );
	BitBlt( hdc, 10, 40, BOMB_LINE*BOMB_WIDTH, BOMB_ROW*BOMB_HEIGHT, mdc, 0, 0, SRCCOPY );
	
	DeleteObject( fullmap );
	DeleteDC( bufdc );
	DeleteDC( mdc );
}

void GetBombPos()
{
	int i;
	for (i=1; i<9; i++ )
	{
		for ( int y=0; y<BOMB_ROW; y++ )
		{
			for ( int x=0; x<BOMB_LINE; x++ )
			{
				if ( BombState[y][x]==i )
				{
					int countB=0;
					int subX = x-1;
					int subY = y-1;
					int addX = x+1;
					int addY = y+1;
					if ( subY>=0 && subX>=0 )
					{
						(BombState[subY][subX]%10==BOMB_UNCLK)?(countB++):(countB);
					}
					if ( subY>=0 )
					{
						(BombState[subY][x]   %10==BOMB_UNCLK)?(countB++):(countB);
					}
					if ( subX>=0 )
					{
						(BombState[y]   [subX]%10==BOMB_UNCLK)?(countB++):(countB);
					}
					if ( addY<BOMB_ROW && subX>=0 )
					{
						(BombState[addY][subX]%10==BOMB_UNCLK)?(countB++):(countB);
					}
					if ( subY>=0 && addX<BOMB_LINE )
					{
						(BombState[subY][addX]%10==BOMB_UNCLK)?(countB++):(countB);
					}
					if ( addX<BOMB_LINE )
					{
						(BombState[y]   [addX]%10==BOMB_UNCLK)?(countB++):(countB);
					}
					if ( addY<BOMB_ROW )
					{
						(BombState[addY][x]   %10==BOMB_UNCLK)?(countB++):(countB);
					}
					if ( addY<BOMB_ROW && addX<BOMB_LINE )
					{
						(BombState[addY][addX]%10==BOMB_UNCLK)?(countB++):(countB);
					}
					if ( BombState[y][x]==countB )
					{
						(subX<0)?(subX=x):(subX);
						(subY<0)?(subY=y):(subY);
						(addX<BOMB_LINE)?(addX):(addX=x);
						(addY<BOMB_ROW )?(addY):(addY=y);

						(BombState[subY][subX]==BOMB_UNCLK)?(BombState[subY][subX]=BOMB_BOMB):(countB);
						(BombState[subY][x]   ==BOMB_UNCLK)?(BombState[subY][x]   =BOMB_BOMB):(countB);
						(BombState[y]   [subX]==BOMB_UNCLK)?(BombState[y]   [subX]=BOMB_BOMB):(countB);
						(BombState[addY][subX]==BOMB_UNCLK)?(BombState[addY][subX]=BOMB_BOMB):(countB);
						(BombState[subY][addX]==BOMB_UNCLK)?(BombState[subY][addX]=BOMB_BOMB):(countB);
						(BombState[y]   [addX]==BOMB_UNCLK)?(BombState[y]   [addX]=BOMB_BOMB):(countB);
						(BombState[addY][x]   ==BOMB_UNCLK)?(BombState[addY][x]   =BOMB_BOMB):(countB);
						(BombState[addY][addX]==BOMB_UNCLK)?(BombState[addY][addX]=BOMB_BOMB):(countB);
						// GetBombPos();
					}				
				}
			}
		}
	}

	for ( i=1; i<9; i++ )
	{
		for ( int y=0; y<BOMB_ROW; y++ )
		{
			for ( int x=0; x<BOMB_LINE; x++ )
			{
				if ( BombState[y][x]==i )
				{
					int countB=0;
					int subX = x-1;
					int subY = y-1;
					int addX = x+1;
					int addY = y+1;
					if ( subY>=0 && subX>=0 )
					{
						(BombState[subY][subX]==BOMB_BOMB)?(countB++):(countB);
					}
					if ( subY>=0 )
					{
						(BombState[subY][x]   ==BOMB_BOMB)?(countB++):(countB);
					}
					if ( subX>=0 )
					{
						(BombState[y]   [subX]==BOMB_BOMB)?(countB++):(countB);
					}
					if ( addY<BOMB_ROW && subX>=0 )
					{
						(BombState[addY][subX]==BOMB_BOMB)?(countB++):(countB);
					}
					if ( subY>=0 && addX<BOMB_LINE )
					{
						(BombState[subY][addX]==BOMB_BOMB)?(countB++):(countB);
					}
					if ( addX<BOMB_LINE )
					{
						(BombState[y]   [addX]==BOMB_BOMB)?(countB++):(countB);
					}
					if ( addY<BOMB_ROW )
					{
						(BombState[addY][x]   ==BOMB_BOMB)?(countB++):(countB);
					}
					if ( addY<BOMB_ROW && addX<BOMB_LINE )
					{
						(BombState[addY][addX]==BOMB_BOMB)?(countB++):(countB);
					}
					if ( BombState[y][x]==countB )
					{
						(subX<0)?(subX=x):(subX);
						(subY<0)?(subY=y):(subY);
						(addX<BOMB_LINE)?(addX):(addX=x);
						(addY<BOMB_ROW )?(addY):(addY=y);

						(BombState[subY][subX]==BOMB_UNCLK)?(BombState[subY][subX]=BOMB_NOBOMB):(countB);
						(BombState[subY][x]   ==BOMB_UNCLK)?(BombState[subY][x]   =BOMB_NOBOMB):(countB);
						(BombState[y]   [subX]==BOMB_UNCLK)?(BombState[y]   [subX]=BOMB_NOBOMB):(countB);
						(BombState[addY][subX]==BOMB_UNCLK)?(BombState[addY][subX]=BOMB_NOBOMB):(countB);
						(BombState[subY][addX]==BOMB_UNCLK)?(BombState[subY][addX]=BOMB_NOBOMB):(countB);
						(BombState[y]   [addX]==BOMB_UNCLK)?(BombState[y]   [addX]=BOMB_NOBOMB):(countB);
						(BombState[addY][x]   ==BOMB_UNCLK)?(BombState[addY][x]   =BOMB_NOBOMB):(countB);
						(BombState[addY][addX]==BOMB_UNCLK)?(BombState[addY][addX]=BOMB_NOBOMB):(countB);
						// GetBombPos();
					}
				}
			}
		}
	}
}

void AIClick()
{
    int kX, kY;
	for ( int y=0; y<BOMB_ROW; y++ )
	{
		for ( int x=0; x<BOMB_LINE; x++ )
		{
			if ( BombState[y][x]==BOMB_NOBOMB )
			{
				// 绝对坐标 = 游戏窗口左上角坐标 + 游戏中的相对坐标
				kX = x*16 + GAME_LEFT + 8;
				kY = y*16 + GAME_TOP + 8;
				LPARAM lParam=(kY<<16)|kX;
				PostMessage( g_WindowHWnd, WM_LBUTTONDOWN, 0, lParam );
				PostMessage( g_WindowHWnd, WM_LBUTTONUP, 0, lParam );
			}
		}
	}
}

void ReDo()
{
	PostMessage( g_WindowHWnd, WM_KEYDOWN, VK_F2, 0 );
}

void RandClick()
{
	time_t t;
	srand((unsigned) time(&t));
	
	int numOfUnclk = 0;
	for ( int y=0; y<BOMB_ROW; y++ )
	{
		for ( int x=0; x<BOMB_LINE; x++ )
		{
			if ( BombState[y][x] == BOMB_UNCLK )
			{
				numOfUnclk++;
			}
		}
	}
	if ( numOfUnclk != 0 )
	{
		numOfUnclk = rand() % numOfUnclk;
	}
	for ( y=0; y<BOMB_ROW; y++ )
	{
		for ( int x=0; x<BOMB_LINE; x++ )
		{
			if ( BombState[y][x] == BOMB_UNCLK )
			{
				numOfUnclk--;
				if ( numOfUnclk == 0 )
				{
					int kX = x*16 + GAME_LEFT + 8;
					int kY = y*16 + GAME_TOP + 8;
					LPARAM lParam=(kY<<16)|kX;
					PostMessage( g_WindowHWnd, WM_LBUTTONDOWN, 0, lParam );
					PostMessage( g_WindowHWnd, WM_LBUTTONUP, 0, lParam );
					return;
				}
			}
		}
	}
}

void SpeedRun()
{
	GameState=GAME_AI;
}

void AIRun()
{
	GameState=GAME_AIRUN;
}

/*
void assert( void* param )
{
	// MessageBox( 0, "请先启动扫雷程序!", "扫雷外挂", 0 );
	// PostQuitMessage(1);
	GameState = GAME_PAUSE;
	SendMessage( hToolsWnd, TB_SETSTATE, (WPARAM)1, (LPARAM)MAKELONG(TBSTATE_INDETERMINATE,0) );
	SendMessage( hToolsWnd, TB_SETSTATE, (WPARAM)2, (LPARAM)MAKELONG(TBSTATE_INDETERMINATE,0) );
}
*/
