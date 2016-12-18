#ifndef COMPONENTS_H_INCLUDED
#define COMPONENTS_H_INCLUDED

#include <windows.h>

void println(const char* str);

//{04175d2a-276b-424c-9ff3-4c0cd65a8373}
const CLSID CLSID_CA = {0x04175d2a,0x276b,0x424c,{0x9f,0xf3,0x4c,0x0c,0xd6,0x5a,0x83,0x73}};

HRESULT __stdcall GetClassObject(const CLSID& clsid, const IID& iid, void** ppv);

#endif // COMPONENTS_H_INCLUDED
