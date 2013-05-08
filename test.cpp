int CTestDlg::DoCommand( LPCTSTR lpszCmdLine )
{
  DWORD ExitCode ;

  try{
    STARTUPINFO StartUpInfo;
    PROCESS_INFORMATION ProcInfo;
    
    ::ZeroMemory( &StartUpInfo, sizeof( StartUpInfo ) ); 

    StartUpInfo.cb            = sizeof( STARTUPINFO ) ; 
    StartUpInfo.dwFlags       = 0 ;
    StartUpInfo.dwFlags       = STARTF_USESHOWWINDOW ;
    StartUpInfo.wShowWindow = SW_SHOW;

    if ( !::CreateProcess( NULL,
                         (LPTSTR)lpszCmdLine,
                         NULL,
                         NULL,
                         FALSE,
                         CREATE_DEFAULT_ERROR_MODE|NORMAL_PRIORITY_CLASS,
                         NULL,
                         NULL,
                         &StartUpInfo,
                         &ProcInfo ) ) {

	    LPVOID lpMsgBuf ;

      ::FormatMessage( FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM,
                       NULL,
                       ::GetLastError(),
                       MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), // Default language
                       (LPTSTR) &lpMsgBuf,
                       0,
                       NULL ) ;

      //// Display the string.
      //AfxMessageBox( (LPCSTR)lpMsgBuf ) ;

      // Free the buffer.
      LocalFree( lpMsgBuf ) ;

      return -1 ;
    } else {

      if( bIsHide ) {
        while ( WaitForSingleObject(ProcInfo.hProcess, 0) != WAIT_OBJECT_0 ) {
          MSG msg;
          while(PeekMessage(&msg,NULL,NULL,NULL,PM_REMOVE)) {
            TranslateMessage(&msg);
            DispatchMessage(&msg);
          }
        }
      } else {
        while ( WaitForSingleObject( ProcInfo.hProcess, INFINITE ) != WAIT_OBJECT_0 );
      }
   
      CloseHandle( ProcInfo.hThread ) ;
    
      ::GetExitCodeProcess( ProcInfo.hProcess, &ExitCode );
      CloseHandle( ProcInfo.hProcess );
    }
  }
  catch(...)
  {
    // do something
  }

  return ExitCode ;
}
