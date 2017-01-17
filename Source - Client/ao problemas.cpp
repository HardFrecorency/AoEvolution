/*
 * Creado: Diciembre 4, 2002
 * Updated: -------------
 * Nombre: AOproblemas
 * Compilado: dev-c++
 * Autor: phiber[imperium]
 * Mail: phiber@clanimperium.com.ar
 * soporte: www.clanimperium.com.ar
 * 
 * USO: solo descompime el zip en una capreta llamada aoproblemas
 * en el mismo directorio donde tiene instalado argentum.
 * ejecuta el exe y listo, problema arreglado.
 *
 */

#include <iostream.h>
#include <stdio.h>
#include <stdlib.h>
#include <conio.c>
#include <process.h> 

		
int main()				
{
system("cls");
textcolor(WHITE);
cout<<"\n";
cout<<"solucionador de AO Runtime errors (v0.1)\n";	
cout<<"----------------------------------------\n";
cout<<"\n";
cout<<"Este programa lo unico que hara en tu sistema es\n";
cout<<"actualizar 3 OCXs (richtx32.ocx, msinet.ocx, cswsk32.ocx)\n";
cout<<"registrarlos y dejar todo listo para el correcto funcionamiento\n";
cout<<"de argentum.\n";
cout<<"\n";
cout<<"\n";
textcolor(YELLOW);
cout<<"COMO USARME: 1ro) Descomprime el zip.\n";
cout<<"             2do) Ve donde tienes instalado ARGENTUM\n";
cout<<"             3ro) Crea una carpeta llamada aoproblemas\n";
cout<<"             4to) Copia todo lo que tenia el zip dentro.\n";
cout<<"             5to) Corre este exe nuevamente.\n";        
cout<<"\n";
cout<<"\n";

textcolor(WHITE);
  int q;	
  char s[1]="";		
cout<<"    *******************************************************\n";
cout<<"    *******************************************************\n";
cout<<"    ***************** desea instalar? (S / N)**************\n"; 				                                           
cout<<"    *******************************************************\n";
cout<<"    *******************************************************\n";   
cout<<"\n";
cout<<"\n";	
cout<<"Respuesta: ";
   cin.getline(s, '\n'); 
    
    q=!strcmpi("s", s);
    
    if (q==1)
  {
 system("cls");
 cout<<"\n";
cout<<"solucionador de AO Runtime errors (v0.1)\n";	
cout<<"----------------------------------------\n";
cout<<"\n";
 textcolor(YELLOW);
 cout<<"continuemos!\n"; 
 textcolor(WHITE);
 cout<<"\n";
  cout<<"**************************************\n";
 cout<<"\n";
 textcolor(WHITE);
system("copy msinet.ocx c:\\windows\\system\\msinet.ocx");
system("copy msinet.ocx c:\\windows\\system32\\msinet.ocx");
  cout<<"\n";
  cout<<"|---> msinet.ocx copiado";
 textcolor(YELLOW);  
cout<<" correctamente!\n";
 cout<<"\n"; 
 textcolor(WHITE);
 cout<<"**************************************\n";  
     
 textcolor(WHITE);
cout<<"\n";
system("copy RICHTX32.OCX ..\\RICHTX32.OCX");
cout<<"\n";
cout<<"|---> ritchx32.ocx copiado";
 textcolor(YELLOW);  
cout<<" correctamente!\n";
 cout<<"\n"; 
 textcolor(WHITE);
 cout<<"**************************************\n"; 
   cout<<"\n";
system("copy CSWSK32.OCX ..\\CSWSK32.OCX");
cout<<"\n";
cout<<"|---> ultimo.ocx copiado";
 textcolor(YELLOW);  
cout<<" correctamente!\n";
 cout<<"\n"; 
 textcolor(WHITE);
 cout<<"\n"; 
 cout<<"**************************************\n";   
     
  system("PAUSE");

 system("cls");
 cout<<"\n";
cout<<"solucionador de AO Runtime errors (v0.1)\n";	
cout<<"----------------------------------------\n";
cout<<"\n";     
       
system("regsvr32.exe /s c:\\windows\\system\\msinet.ocx");
system("regsvr32.exe /s c:\\windows\\system32\\msinet.ocx"); 
cout<<"\n";
  cout<<"|---> msinet.ocx registrado";
 textcolor(YELLOW);  
cout<<" correctamente!\n";
 cout<<"\n"; 
 textcolor(WHITE);
 cout<<"**************************************\n";  

cout<<"\n";     
       
system("regsvr32.exe /s ..\\RICHTX32.OCX");
 
cout<<"\n";
  cout<<"|---> ritchtx32.ocx registrado";
 textcolor(YELLOW);  
cout<<" correctamente!\n";
 cout<<"\n"; 
 textcolor(WHITE);
 cout<<"**************************************\n";          
         
  cout<<"\n";     
       
system("regsvr32.exe /s ..\\CSWSK32.OCX");
 
cout<<"\n";
  cout<<"|---> cswsk32.ocx registrado";
 textcolor(YELLOW);  
cout<<" correctamente!\n";
 cout<<"\n"; 
 textcolor(WHITE);
 cout<<"**************************************\n";          
 system("PAUSE");          
system("cls");
cout<<"\n";
cout<<"solucionador de AO Runtime errors (v0.1)\n";	
cout<<"----------------------------------------\n";
cout<<"\n";
textcolor(YELLOW);
cout<<"TERMINADO!\n";
cout<<"\n";
textcolor(WHITE);
cout<<"Ahora solo te queda probar argentum.\n";
cout<<"Gracias por usar nuestro software.\n";
cout<<"\n";
cout<<"\n";
cout<<"\n";
cout<<"\n";
textcolor(YELLOW);
cout<<"Phiber[imperium]\n";
cout<<"www.clanimperium.com.ar\n";
textcolor(WHITE);
system("PAUSE");
                     
  }


  else 		
   
  {
    cout<<"";
    cout<<"";
    cout<<"";
    cout<<"";
    textcolor(YELLOW);		
    cout<<"bye bye";
    textcolor(WHITE);
  }

  return 0;
}  
