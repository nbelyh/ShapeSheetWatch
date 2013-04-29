
#include "stdafx.h"
#include "Addin.h"

#include <GdiPlus.h>
#pragma comment(lib,"gdiplus.lib")

HRESULT CustomUiGetPng(LPCWSTR resource_id, IPictureDisp ** result_image, IPictureDisp** result_mask)
{
	HMODULE hModule = AfxGetResourceHandle();
	HRESULT hr = S_OK;

	using namespace Gdiplus;

	PICTDESC pd = {0};
	pd.cbSizeofstruct = sizeof (PICTDESC);
	pd.picType = PICTYPE_BITMAP;

	GdiplusStartupInput gdiplusStartupInput;
	ULONG_PTR gdiplusToken;

	gdiplusStartupInput.DebugEventCallback = NULL;
	gdiplusStartupInput.SuppressBackgroundThread = FALSE;
	gdiplusStartupInput.SuppressExternalCodecs = FALSE;
	gdiplusStartupInput.GdiplusVersion = 1;
	GdiplusStartup (&gdiplusToken, &gdiplusStartupInput, NULL);

	HRSRC hResource = FindResource (hModule, resource_id, L"PNG");

	if (!hResource)
		return HRESULT_FROM_WIN32(GetLastError());

	DWORD dwImageSize = SizeofResource (hModule, hResource);

	const void* pResourceData = LockResource (LoadResource(hModule, hResource));
	if (!pResourceData)
		return HRESULT_FROM_WIN32(GetLastError());

	HGLOBAL hBuffer = GlobalAlloc (GMEM_MOVEABLE, dwImageSize);
	if (hBuffer)
	{
		void* pBuffer = GlobalLock (hBuffer);
		if (pBuffer)
		{
			CopyMemory (pBuffer, pResourceData, dwImageSize);

			IStream* pStream = NULL;
			if (::CreateStreamOnHGlobal (hBuffer, FALSE, &pStream) == S_OK)
			{
				Bitmap *bitmap = Bitmap::FromStream (pStream);
				pStream->Release();

				if (bitmap)
				{
					if (result_mask == NULL) // direct support for picture
					{
						bitmap->GetHBITMAP (0, &pd.bmp.hbitmap);
						hr = OleCreatePictureIndirect (&pd, IID_IDispatch, TRUE, (LPVOID*)result_image);
					}
					else	// old version - 2003/2007 - split into picture + mask
					{
						UINT w = bitmap->GetWidth();
						UINT h = bitmap->GetHeight();
						Rect r(0, 0, w, h);

						Bitmap bitmap_rgb(w, h, PixelFormat24bppRGB);
						Bitmap bitmap_msk(w, h, PixelFormat24bppRGB);

						BitmapData argb_bits;
						bitmap->LockBits(&r, ImageLockModeRead, PixelFormat32bppARGB, &argb_bits);

						BitmapData rgb_bits;
						bitmap_rgb.LockBits(&r, ImageLockModeWrite, PixelFormat24bppRGB, &rgb_bits);

						BitmapData msk_bits;
						bitmap_msk.LockBits(&r, ImageLockModeWrite, PixelFormat24bppRGB, &msk_bits);

						for (UINT y = 0; y < h; ++y)
						{
							for (UINT x = 0; x < w; ++x)
							{
								BYTE* idx_argb = 
									static_cast<BYTE*>(argb_bits.Scan0) + (x + y*w) * 4;

								BYTE* idx_rgb = 
									static_cast<BYTE*>(static_cast<BYTE*>(rgb_bits.Scan0) + (x + y*w) * 3);

								BYTE* idx_msk =
									static_cast<BYTE*>(static_cast<BYTE*>(msk_bits.Scan0) + (x + y*w) * 3);


								idx_rgb[0] = idx_argb[0];
								idx_rgb[1] = idx_argb[1];
								idx_rgb[2] = idx_argb[2];

								byte t = (idx_argb[3] < 128) ? 0xFF : 0x00;

								idx_msk[0] = t;
								idx_msk[1] = t;
								idx_msk[2] = t;
							}
						}

						bitmap->UnlockBits(&argb_bits);

						bitmap_rgb.UnlockBits(&rgb_bits);
						bitmap_msk.UnlockBits(&msk_bits);

						bitmap_rgb.GetHBITMAP (0, &pd.bmp.hbitmap);
						hr = OleCreatePictureIndirect (&pd, IID_IDispatch, TRUE, (LPVOID*)result_image);

						bitmap_msk.GetHBITMAP (0, &pd.bmp.hbitmap);
						hr = OleCreatePictureIndirect (&pd, IID_IDispatch, TRUE, (LPVOID*)result_mask);
					}


					delete bitmap;
				}
			}
			GlobalUnlock (pBuffer);
		}
		GlobalFree (hBuffer);
	}

	GdiplusShutdown (gdiplusToken);
	return hr;
}
