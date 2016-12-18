#include "interfaces.h"
#include <stdio.h>
#include "Components.h"
#include "InClass.h"
#include <fstream>
#include <iostream>
#include <string>
using namespace std;

//*************************************************

class CA: public IInput, IOutput, IInterpolation, IDispatch
{
   private:
       struct {
           vector< vector<double> > result;
           int units;
       } resultInterBCube;
    int counter;
    int px1;
    int *values;
    int units;
    string filename;
    vector< vector<double> > resBLin;
   public:
    CA();
    virtual ~CA();

    virtual HRESULT __stdcall QueryInterface(const IID& iid, void** ppv);
	virtual ULONG __stdcall AddRef();
	virtual ULONG __stdcall Release();

	virtual void __stdcall InputFromCmd(int *values);
	virtual void __stdcall InputFromFile(string filename);

	virtual void __stdcall OutputToCmd(vector< vector<double> > result, int units);
	virtual void __stdcall OutputToFile();

	virtual void __stdcall InterpolationTriple(int *values);
	virtual void __stdcall InterpolationBCube(int *values);

	virtual HRESULT __stdcall GetTypeInfoCount(UINT* pctinfo);
	virtual HRESULT __stdcall GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo** ppTInfo);
	virtual HRESULT __stdcall GetIDsOfNames(REFIID riid, LPOLESTR* rgszNames, UINT cNames, LCID lcid, DISPID* rgDispId);
	virtual HRESULT __stdcall Invoke(DISPID dispIdMember, REFIID riid, LCID lcid,WORD wFlags, DISPPARAMS* pDispParams,VARIANT* pVarResult,
                                     EXCEPINFO* pExcepInfo, UINT* puArgErr);

};

CA::CA()
{
  println("CA::CA");
  counter = 0;
  px1=123;
}

CA::~CA()
{
  println("CA::~CA");
}


HRESULT __stdcall CA::QueryInterface(const IID& iid, void** ppv)
{
   println("CA::QueryInterface");

   if (iid==IID_IUnknown)
   {
      *ppv = (IUnknown*)(IInput*)this;
   }
   else if (iid==IID_IInput)
   {
      *ppv = (IInput*)this;
   }
   else if (iid==IID_IOutput)
   {
      *ppv = (IOutput*)this;
   }
   else if (iid==IID_IInterpolation)
   {
      *ppv = (IInterpolation*)this;
   }
   else if (iid==IID_IDispatch)
   {
      *ppv = (IDispatch*)this;
   }
   else
   {
     *ppv = NULL;
     return E_NOINTERFACE;
   }

   this->AddRef();
   return S_OK;
}

ULONG __stdcall CA::AddRef()
{
   println("CA::AddRef");
   counter = counter + 1;
   return counter;
}

ULONG __stdcall CA::Release()
{
   println("CA::Release");
   counter = counter - 1;
   if (counter==0) {delete this;}
   return counter;
}

void __stdcall CA::InputFromCmd(int *values)
{
   println("CA:InputFromCmd");
   this.values = values;
}

void __stdcall CA::InputFromFile(string filename)
{
   println("CA:InputFromFile");
   ifstream fin("in.txt");

   int num = 0;

   while (!fin.eof()) {
        fin >> num;
        values = new int[num];
        for (int i = 0; i < num; i++) {
            fin >> values[i];
        }
        fin.close();
   }

   InterpolationBCube(values);
}

void __stdcall CA::OutputToCmd(vector< vector<double> > result, int units)
{
   println("CA:OutputToCmd");
   /*for (int i = 0; i < 5; i++) {
        cout << values[i] << endl;
   }*/
   for (int i = 0; i < units; i ++){
        for(int j = 0; j < units; j ++)
            cout << result[i][j] << "   ";
        cout << endl;
   }
}

void __stdcall CA::OutputToFile()
{
   println("CA:OutputToFile");
}

void __stdcall CA::InterpolationTriple(int *values)
{
   println("CA:InterpolationTriple");
   for (int i = 0; i < 5; i++) {
        values[i] += i;
   }
   //OutputToCmd(values);
}

void __stdcall CA::InterpolationBCube(int *values)
{
    println("CA:InterpolationBCube");
    int units = values[4];
    InterBi inclass(values[0], values[1], values[2], values[3], units);
    vector< vector<double> > result;
        result.resize(units);
        for(int i = 0; i < units; i ++){
            result[i].resize(units);
        }
    result = inclass.interpol(inclass.getX(), inclass.getY());
    resultInterBCube.result = result;
    resultInterBCube.units = units;
}

//IDispatch (Begin)
HRESULT __stdcall CA::GetTypeInfoCount(UINT* pctinfo)
{
    println("CA:GetTypeInfoCount");
    return S_OK;
}

HRESULT __stdcall CA::GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo** ppTInfo)
{
    println("CA:GetTypeInfo");
    return S_OK;
}

HRESULT __stdcall CA::GetIDsOfNames(REFIID riid, LPOLESTR* rgszNames, UINT cNames,
                                    LCID lcid, DISPID* rgDispId)
{
    println("CA:GetIDsOfNames");
    if (cNames!=1) {return E_NOTIMPL;}

    if (wcscmp(rgszNames[0],L"InputFromCmd")==0)
    {
      rgDispId[0] = 1;
    }
    else if (wcscmp(rgszNames[0],L"InputFromFile")==0)
    {
      rgDispId[0] = 2;
    }
    else if (wcscmp(rgszNames[0],L"OutputToCmd")==0)
    {
      rgDispId[0] = 3;
    }
    else if (wcscmp(rgszNames[0],L"OutputToFile")==0)
    {
      rgDispId[0] = 4;
    }
    else if (wcscmp(rgszNames[0],L"InterpolationTriple")==0)
    {
      rgDispId[0] = 5;
    }
    else if (wcscmp(rgszNames[0],L"InterpolationBCube")==0)
    {
      rgDispId[0] = 6;
    }


    //Property
    else if (wcscmp(rgszNames[0],L"PInputFromCmd")==0)
    {
      rgDispId[0] = 7;
    }
    else
    {
       return E_NOTIMPL;
    }
    return S_OK;
}

HRESULT __stdcall CA::Invoke(DISPID dispIdMember, REFIID riid, LCID lcid,WORD wFlags, DISPPARAMS* pDispParams,VARIANT* pVarResult,
                             EXCEPINFO* pExcepInfo, UINT* puArgErr)
{
    println("CA:Invoke");
    if (dispIdMember==1)
    {
       InputFromCmd(values);
    }
    else if (dispIdMember==2)
    {
       InputFromFile(filename);
    }
    else if (dispIdMember==3)
    {
       OutputToCmd(resBLin, units);
    }
    else if (dispIdMember==4)
    {
       OutputToFile();
    }
    else if (dispIdMember==5)
    {
       InterpolationTriple(values);
    }
    else if (dispIdMember==6)
    {
       InterpolationBCube(values);
    }

    //Property
    else if (dispIdMember==7)
    {
       if (wFlags==DISPATCH_PROPERTYGET)
       {
          pVarResult->vt = VT_INT;
          pVarResult->intVal = px1;
       }
       else if (wFlags==DISPATCH_PROPERTYPUT)
       {
          DISPPARAMS param = *pDispParams;
          VARIANT arg = (param.rgvarg)[0];
          px1 = arg.intVal;
       }
       else
       {
         return E_NOTIMPL;
       }
    }
    else
    {
      return E_NOTIMPL;
    }
    return S_OK;
}
//IDispatch (End)


//*************************************************


class CFA: public IClassFactory
{
    private:
     int counter;
    public:
      CFA();
      virtual ~CFA();

    virtual HRESULT __stdcall QueryInterface(const IID& iid, void** ppv);
	virtual ULONG __stdcall AddRef();
	virtual ULONG __stdcall Release();

	virtual HRESULT __stdcall CreateInstance(IUnknown* pUnknownOuter, const IID& iid, void** ppv);
	virtual HRESULT __stdcall LockServer(BOOL bLock);
};

CFA::CFA()
{
  println("CFA::CFA");
  counter = 0;
}

CFA::~CFA()
{
  println("CFA::~CFA");
}


HRESULT __stdcall CFA::QueryInterface(const IID& iid, void** ppv)
{
   println("CFA::QueryInterface");

   if (iid==IID_IUnknown)
   {
     *ppv = (IUnknown*)(IClassFactory*)this;
   }
   else if (iid==IID_IClassFactory)
   {
     *ppv = (IClassFactory*)this;
   }
   else
   {
     *ppv = NULL;
     return E_NOINTERFACE;
   }
   this->AddRef();
   return S_OK;
}

ULONG __stdcall CFA::AddRef()
{
   println("CFA::AddRef");
   counter = counter + 1;
   return counter;
}

ULONG __stdcall CFA::Release()
{
   println("CFA::Release");
   counter = counter - 1;
   if (counter==0) {delete this;}
   return counter;
}


HRESULT __stdcall CFA::CreateInstance(IUnknown* pUnknownOuter, const IID& iid, void** ppv)
{
   println("CFA::CreateInstance");
   if (pUnknownOuter!=NULL)
   {
      return E_NOTIMPL;
   }

   CA* a = new CA();
   return a->QueryInterface(iid,ppv);
}

HRESULT __stdcall CFA::LockServer(BOOL bLock)
{
  println("CFA::LockServer");
  return S_OK;
}

//*************************************************


void println(const char* str)
{
    printf(str);
    printf("\n");
}

HRESULT __stdcall GetClassObject(const CLSID& clsid, const IID& iid, void** ppv)
{
   println("Component::GetClassObject");
   if (clsid==CLSID_CA)
   {
      CFA* fa  = new CFA();
      return fa->QueryInterface(iid,ppv);
   }
   else
   {
     *ppv = NULL;
     return E_NOTIMPL;
   }

}

//*************************************************
