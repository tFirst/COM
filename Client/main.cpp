#include <objbase.h>
#include <iostream>
#include <stdio.h>
#include <locale>

#include "Interfaces.h"
using namespace std;

//{04175d2a-276b-424c-9ff3-4c0cd65a8373}
const CLSID CLSID_CA = {0x04175d2a,0x276b,0x424c,{0x9f,0xf3,0x4c,0x0c,0xd6,0x5a,0x83,0x73}};

//Manager emulator method
typedef HRESULT __stdcall (*GetClassObjectType) (const CLSID& clsid, const IID& iid, void** ppv);


int main()
{
    setlocale(LC_ALL, "");
    printf("Main::Start\n");

    int var = 0;
    //var = 1; //Using manager emulator
    //var = 2; //Using COM library (factory)
    var = 3; //Using COM library (instance)


    //Initializing COM library (Begin)
    CoInitialize(NULL);
    //Initializing COM library (End)
    int type;
    cout << "Enter type of input(1 or 2): " << endl;
    cin >> type;


    try
    {
     CLSID CLSID_CA_ProgID;
     bool flagProgID = true;
     {
      const wchar_t* progID = L"Trushin.Interpolation";
      //mbstowcs
      //wcstombs
      HRESULT resCLSID_ProgID = CLSIDFromProgID(progID,&CLSID_CA_ProgID);
      if (!(SUCCEEDED(resCLSID_ProgID)))
      {
        flagProgID = false;
        printf("No CLSID form ProgID");
        printf("\n");
      }
      else
      {
        printf("CLSID form ProgID OK!");
        printf("\n");
      }
     }

     IInput* pInput = NULL;
     IClassFactory* pCF = NULL;
     int values[5] = {1, 2, 5, 6, 5};

     if (var!=3)
     {

      HRESULT resFactory;

      if (var==1)
      {
        //Getting manager method (Begin)
        GetClassObjectType GetClassObject;
        HINSTANCE h;
        h=LoadLibrary("../../../Manager/bin/Release/Manager.dll");
        if (!h)
        {
          throw "No manager";
        }
        GetClassObject=(GetClassObjectType) GetProcAddress(h,"GetClassObject");
        if (!GetClassObject)
        {
          throw "No manager method";
        }
        //Getting manager method (End)

        //Getting factory (Begin)
        if (flagProgID==true)
        {
          resFactory = GetClassObject(CLSID_CA_ProgID,IID_IClassFactory,(void**)&pCF);
        }
        else
        {
          resFactory = GetClassObject(CLSID_CA,IID_IClassFactory,(void**)&pCF);
        }

        //Getting factory (End)
      }
      else
      {
        //Getting factory (Begin)
        if (flagProgID==true)
        {
          resFactory = CoGetClassObject(CLSID_CA_ProgID,CLSCTX_INPROC_SERVER,NULL,IID_IClassFactory,(void**)&pCF);
        }
        else
        {
          resFactory = CoGetClassObject(CLSID_CA,CLSCTX_INPROC_SERVER,NULL,IID_IClassFactory,(void**)&pCF);
        }
        //Getting factory (End)
      }

      if (!(SUCCEEDED(resFactory)))
      {
         //printf("%X\n",(unsigned int)resFactory);
         throw "No factoty";
      }

     } // if need factory


     HRESULT resInstance;

     //Getting instance (Begin)
     if (var!=3)
     {
       resInstance = pCF->CreateInstance(NULL,IID_IInput,(void**)&pInput);
     }
     else
     {
       resInstance = CoCreateInstance(CLSID_CA,NULL,CLSCTX_INPROC_SERVER,IID_IInput,(void**)&pInput);
     }

     if (!(SUCCEEDED(resInstance)))
     {
          //printf("%X\n",(unsigned int)resInstance);
          throw "No instance";
     }
     //Getting instance (End)


     //Work (Begin)
     if (type == 1) {
        pInput->InputFromCmd(values);
     } else {
         pInput->InputFromFile("in.txt");
     }
     //Work (End)

     IOutput* pOutput = NULL;
     HRESULT resQuery = pInput->QueryInterface(IID_IOutput,(void**)&pOutput);
     if (!(SUCCEEDED(resQuery)))
     {
          //printf("%X\n",(unsigned int)resQuery);
          throw "No query";
     }
     //Query (End)

     //Work (Begin)
     //pOutput->OutputToCmd(values);
     //Work (End)

     //Dispatch (Begin)
     IDispatch* pDisp = NULL;
     HRESULT resQueryDisp = pInput->QueryInterface(IID_IDispatch,(void**)&pDisp);
     if (!(SUCCEEDED(resQueryDisp)))
     {
          //printf("%X\n",(unsigned int)resQuery);
          throw "No query dispatch";
     }

     DISPID dispid;

     int namesCount = 1;
     OLECHAR** names = new OLECHAR*[namesCount];
     OLECHAR* name = const_cast<OLECHAR*>(L"OutputToCmd");
     names[0] = name;
     HRESULT resID_Name = pDisp->GetIDsOfNames(
                                                    IID_NULL, // Должно быть IID_NULL
                                                    names, // Имя функции
                                                    namesCount, // Число имен
                                                    GetUserDefaultLCID(), // Информация локализации
                                                    &dispid
                                              );
     if (!(SUCCEEDED(resID_Name)))
     {
          //printf("%X\n",(unsigned int)resID_Name);
          printf("No ID of name\n");
     }
     else
     {
         DISPPARAMS dispparamsNoArgs = {
                                         NULL,
                                         NULL,
                                         0, // Ноль аргументов
                                         0, // Ноль именованных аргументов
                                       };

         HRESULT resInvoke = pDisp->Invoke(
                                            dispid, // DISPID
                                            IID_NULL, // Должно быть IID_NULL
                                            GetUserDefaultLCID(), // Информация локализации
                                            DISPATCH_METHOD, // Метод
                                            &dispparamsNoArgs, // Аргументы метода
                                            NULL, // Результаты
                                            NULL, // Исключение
                                            NULL
                                          ); // Ошибка в аргументе
        if (!(SUCCEEDED(resInvoke)))
        {
          //printf("%X\n",(unsigned int)resInvoke);
          printf("Invoke error\n");
        }
     }

     pDisp->Release();
     //Dispatch (End)

     //Deleting (Begin)
     pOutput->Release();

     //Query (Begin)

     pInput->Release();
     if (var!=3)
     {
       pCF->Release();
     }
     //Deleting (End)

    }
    catch (const char* str)
    {
        printf("Main::Error: ");
        printf(str);
        printf("\n");
    }
    catch (...)
    {
        printf("Main::Error: Unknown\n");
    }


    //Uninitializing COM library (Begin)
    CoUninitialize();
    //Uninitializing COM library (End)

    printf("Main::Finish\n");
    return 0;
}
