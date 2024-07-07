#include <algorithm>
#include <cstdint>
#include <new>
#include <shlwapi.h>
#include <string>
#include <thumbcache.h> /* for #IThumbnailProvider */
#include <vector>

#include "Wincodec.h"

#pragma comment(lib, "shlwapi.lib")

struct Thumbnail {
  std::vector<uint8_t> data;
  int width;
  int height;
};

/**
 * This thumbnail provider implements #IInitializeWithStream to enable being
 * hosted in an isolated process for robustness.
 */
class CKisekiThumb : public IInitializeWithStream, public IThumbnailProvider {
public:
  CKisekiThumb() : _cRef(1), _pStream(nullptr) {}

  virtual ~CKisekiThumb() {
    if (_pStream) {
      _pStream->Release();
    }
  }

  IFACEMETHODIMP QueryInterface(REFIID riid, void **ppv) {
    static const QITAB qit[] = {
        QITABENT(CKisekiThumb, IInitializeWithStream),
        QITABENT(CKisekiThumb, IThumbnailProvider),
        {0},
    };
    return QISearch(this, qit, riid, ppv);
  }

  IFACEMETHODIMP_(ULONG) AddRef() { return InterlockedIncrement(&_cRef); }

  IFACEMETHODIMP_(ULONG) Release() {
    ULONG cRef = InterlockedDecrement(&_cRef);
    if (!cRef) {
      delete this;
    }
    return cRef;
  }

  /** IInitializeWithStream */
  IFACEMETHODIMP Initialize(IStream *pStream, DWORD grfMode);

  /** IThumbnailProvider */
  IFACEMETHODIMP GetThumbnail(UINT cx, HBITMAP *phbmp, WTS_ALPHATYPE *pdwAlpha);

private:
  long _cRef;
  IStream *_pStream; /* provided in Initialize(). */
};

HRESULT CKisekiThumb_CreateInstance(REFIID riid, void **ppv) {
  CKisekiThumb *pNew = new (std::nothrow) CKisekiThumb();
  HRESULT hr = pNew ? S_OK : E_OUTOFMEMORY;
  if (SUCCEEDED(hr)) {
    hr = pNew->QueryInterface(riid, ppv);
    pNew->Release();
  }
  return hr;
}

IFACEMETHODIMP CKisekiThumb::Initialize(IStream *pStream, DWORD) {
  if (_pStream != nullptr) {
    /* Can only be initialized once. */
    return E_UNEXPECTED;
  }
  /* Take a reference to the stream. */
  return pStream->QueryInterface(&_pStream);
}

IFACEMETHODIMP CKisekiThumb::GetThumbnail(UINT cx, HBITMAP *phbmp,
                                          WTS_ALPHATYPE *pdwAlpha) {
  HRESULT hr = S_FALSE;

  std::vector<char> buffer;
  ULONG bytesRead;
  char chunk[4096];
  do {
    hr = _pStream->Read(chunk, sizeof(chunk), &bytesRead);
    if (SUCCEEDED(hr) && bytesRead > 0) {
      buffer.insert(buffer.end(), chunk, chunk + bytesRead);
    }
  } while (SUCCEEDED(hr) && bytesRead > 0);

  if (FAILED(hr)) {
    return hr;
  }

  // Find the </roblox> closing tag
  const std::string endTag = "</roblox>";
  auto it =
      std::search(buffer.begin(), buffer.end(), endTag.begin(), endTag.end());
  if (it == buffer.end()) {
    return E_FAIL; // Closing tag not found
  }

  // Move iterator past the closing tag and the null byte
  std::advance(it, endTag.length() + 1);

  if (it >= buffer.end()) {
    return E_FAIL; // No data after the closing tag
  }

  // The rest is JPEG data
  std::vector<unsigned char> jpegData(it, buffer.end());

  // Create a WIC factory
  IWICImagingFactory *pFactory = nullptr;
  hr = CoCreateInstance(CLSID_WICImagingFactory, nullptr, CLSCTX_INPROC_SERVER,
                        IID_PPV_ARGS(&pFactory));
  if (FAILED(hr)) {
    return hr;
  }

  // Create a stream from the JPEG data
  IWICStream *pStream = nullptr;
  hr = pFactory->CreateStream(&pStream);
  if (FAILED(hr)) {
    pFactory->Release();
    return hr;
  }

  hr = pStream->InitializeFromMemory(jpegData.data(), jpegData.size());
  if (FAILED(hr)) {
    pStream->Release();
    pFactory->Release();
    return hr;
  }

  // Create a decoder
  IWICBitmapDecoder *pDecoder = nullptr;
  hr = pFactory->CreateDecoderFromStream(
      pStream, nullptr, WICDecodeMetadataCacheOnDemand, &pDecoder);
  if (FAILED(hr)) {
    pStream->Release();
    pFactory->Release();
    return hr;
  }

  // Get the first frame of the image from the decoder
  IWICBitmapFrameDecode *pFrame = nullptr;
  hr = pDecoder->GetFrame(0, &pFrame);
  if (FAILED(hr)) {
    pDecoder->Release();
    pStream->Release();
    pFactory->Release();
    return hr;
  }

  // Get the size of the image
  UINT width, height;
  hr = pFrame->GetSize(&width, &height);
  if (FAILED(hr)) {
    pFrame->Release();
    pDecoder->Release();
    pStream->Release();
    pFactory->Release();
    return hr;
  }

  UINT stride = (width * 3 + 3) & ~3;

  // Create a bitmap and copy the pixels
  Thumbnail thumb;
  thumb.width = width;
  thumb.height = height;
  thumb.data.resize(height * stride);
  hr =
      pFrame->CopyPixels(nullptr, stride, thumb.data.size(), thumb.data.data());

  pFrame->Release();
  pDecoder->Release();
  pStream->Release();
  pFactory->Release();

  if (FAILED(hr)) {
    return hr;
  }

  *phbmp = CreateBitmap(thumb.width, thumb.height, 1, 24, thumb.data.data());
  if (!*phbmp) {
    return E_FAIL;
  }
  *pdwAlpha = WTSAT_RGB;

  hr = S_OK;
  return hr;
}
