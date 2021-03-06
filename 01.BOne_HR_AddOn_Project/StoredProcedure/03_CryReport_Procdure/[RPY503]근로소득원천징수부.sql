IF(EXISTS(SELECT NAME FROM sysobjects WHERE NAME = 'RPY503' AND xtype = 'P'))
	DROP PROCEDURE RPY503
GO

CREATE  PROC RPY503
    (
        @JSNYER     AS Nvarchar(4),     --작업연월
        @STRMON     AS Nvarchar(2),     --시작월
        @ENDMON     AS Nvarchar(2),     --종료월
        @JOBGBN     AS Nvarchar(1),     --작업구분(1연말정산,2중도정산,3전체)
        @CLTCOD     AS Nvarchar(8),     --자사코드
        @MSTDPT     AS Nvarchar(8),     --부서
        @MSTCOD     AS Nvarchar(8)      --사원번호      
    )
    

 AS
 
    /*==========================================================================================
        프로시저명      : RPY503
        프로시저설명    : 근로소득원천징수부
        만든이          : 함미경
        작업일자        : 2007-01-30
        작업지시자      : 함미경
        작업지시일자    : 2007-01-30
        작업목적        : 
        작업내용        : 
    ===========================================================================================*/
    --DROP PROC RPY503
    --Exec RPY503 '2009', '01', '12', '2', N'%', N'%', N'%', N'%'

    SET NOCOUNT ON
    
    CREATE TABLE #RPY503 (
            MSTCOD     nvarchar(8) COLLATE Korean_Wansung_CI_AS,
            MSTNAM     nvarchar(50),
            STRDAY	   Datetime,
			ENDDAY	   Datetime,
            PERNBR     nvarchar(20),
            BUYNSU     Numeric(19,6),   

            JIGM01      nvarchar(6),    
            PAYM01      Numeric(19,6),  
            BNSM01      Numeric(19,6),  
            INJB01      Numeric(19,6),  
            SUMM01      Numeric(19,6),  
            BT1M01      Numeric(19,6),  
            BT2M01      Numeric(19,6),  
            BT3M01      Numeric(19,6),      
            GABM01      Numeric(19,6),      
            JUMM01      Numeric(19,6),  
            KUKM01      Numeric(19,6),  
            MEDM01      Numeric(19,6),  
            GBHM01      Numeric(19,6),
            JUSM01      Numeric(19,6),
            GBUM01      Numeric(19,6),
            BU3M01      Numeric(19,6),
            BT4M01      Numeric(19,6),
            BT5M01      Numeric(19,6),
            BT6M01      Numeric(19,6),
            JIGM02      nvarchar(6),    
            PAYM02      Numeric(19,6),  
            BNSM02      Numeric(19,6),  
            INJB02      Numeric(19,6),  
            SUMM02      Numeric(19,6),  
            BT1M02      Numeric(19,6),  
            BT2M02      Numeric(19,6),  
            BT3M02      Numeric(19,6),      
            GABM02      Numeric(19,6),      
            JUMM02      Numeric(19,6),  
            KUKM02      Numeric(19,6),  
            MEDM02      Numeric(19,6),  
            GBHM02      Numeric(19,6),
            JUSM02      Numeric(19,6),
            GBUM02      Numeric(19,6),
            BU3M02      Numeric(19,6),
            BT4M02      Numeric(19,6),
            BT5M02      Numeric(19,6),
            BT6M02      Numeric(19,6),
            JIGM03      nvarchar(6),    
            PAYM03      Numeric(19,6),  
            BNSM03      Numeric(19,6),  
            INJB03      Numeric(19,6),  
            SUMM03      Numeric(19,6),  
            BT1M03      Numeric(19,6),  
            BT2M03      Numeric(19,6),  
            BT3M03      Numeric(19,6),      
            GABM03      Numeric(19,6),      
            JUMM03      Numeric(19,6),  
            KUKM03      Numeric(19,6),  
            MEDM03      Numeric(19,6),  
            GBHM03      Numeric(19,6),
            JUSM03      Numeric(19,6),
            GBUM03      Numeric(19,6),
            BU3M03      Numeric(19,6),
            BT4M03      Numeric(19,6),          
            BT5M03      Numeric(19,6),
            BT6M03      Numeric(19,6),
            JIGM04      nvarchar(6),    
            PAYM04      Numeric(19,6),  
            BNSM04      Numeric(19,6),  
            INJB04      Numeric(19,6),  
            SUMM04      Numeric(19,6),  
            BT1M04      Numeric(19,6),  
            BT2M04      Numeric(19,6),  
            BT3M04      Numeric(19,6),      
            GABM04      Numeric(19,6),      
            JUMM04      Numeric(19,6),  
            KUKM04      Numeric(19,6),  
            MEDM04      Numeric(19,6),  
            GBHM04      Numeric(19,6),
            JUSM04      Numeric(19,6),
            GBUM04      Numeric(19,6),
            BU3M04      Numeric(19,6),
            BT4M04      Numeric(19,6),
            BT5M04      Numeric(19,6),
            BT6M04      Numeric(19,6),
            JIGM05      nvarchar(6),    
            PAYM05      Numeric(19,6),  
            BNSM05      Numeric(19,6),  
            INJB05      Numeric(19,6),  
            SUMM05      Numeric(19,6),  
            BT1M05      Numeric(19,6),  
            BT2M05      Numeric(19,6),  
            BT3M05      Numeric(19,6),      
            GABM05      Numeric(19,6),      
            JUMM05      Numeric(19,6),  
            KUKM05      Numeric(19,6),  
            MEDM05      Numeric(19,6),  
            GBHM05      Numeric(19,6),
            JUSM05      Numeric(19,6),
            GBUM05      Numeric(19,6),
            BU3M05      Numeric(19,6),
            BT4M05      Numeric(19,6),
            BT5M05      Numeric(19,6),
            BT6M05      Numeric(19,6),          
            JIGM06      nvarchar(6),    
            PAYM06      Numeric(19,6),  
            BNSM06      Numeric(19,6),  
            INJB06      Numeric(19,6),  
            SUMM06      Numeric(19,6),  
            BT1M06      Numeric(19,6),  
            BT2M06      Numeric(19,6),  
            BT3M06      Numeric(19,6),      
            GABM06      Numeric(19,6),      
            JUMM06      Numeric(19,6),  
            KUKM06      Numeric(19,6),  
            MEDM06      Numeric(19,6),  
            GBHM06      Numeric(19,6),
            JUSM06      Numeric(19,6),
            GBUM06      Numeric(19,6),
            BU3M06      Numeric(19,6),
            BT4M06      Numeric(19,6),
            BT5M06      Numeric(19,6),
            BT6M06      Numeric(19,6),
            JIGM07      nvarchar(6),    
            PAYM07      Numeric(19,6),  
            BNSM07      Numeric(19,6),  
            INJB07      Numeric(19,6),  
            SUMM07      Numeric(19,6),  
            BT1M07      Numeric(19,6),  
            BT2M07      Numeric(19,6),  
            BT3M07      Numeric(19,6),      
            GABM07      Numeric(19,6),      
            JUMM07      Numeric(19,6),  
            KUKM07      Numeric(19,6),  
            MEDM07      Numeric(19,6),  
            GBHM07      Numeric(19,6),
            JUSM07      Numeric(19,6),
            GBUM07      Numeric(19,6),
            BU3M07      Numeric(19,6),
            BT4M07      Numeric(19,6),
            BT5M07      Numeric(19,6),
            BT6M07      Numeric(19,6),          
            JIGM08      nvarchar(6),    
            PAYM08      Numeric(19,6),  
            BNSM08      Numeric(19,6),  
            INJB08      Numeric(19,6),  
            SUMM08      Numeric(19,6),  
            BT1M08      Numeric(19,6),  
            BT2M08      Numeric(19,6),  
            BT3M08      Numeric(19,6),      
            GABM08      Numeric(19,6),      
            JUMM08      Numeric(19,6),  
            KUKM08      Numeric(19,6),  
            MEDM08      Numeric(19,6),  
            GBHM08      Numeric(19,6),
            JUSM08      Numeric(19,6),
            GBUM08      Numeric(19,6),
            BU3M08      Numeric(19,6),
            BT4M08      Numeric(19,6),
            BT5M08      Numeric(19,6),
            BT6M08      Numeric(19,6),          
            JIGM09      nvarchar(6),    
            PAYM09      Numeric(19,6),  
            BNSM09      Numeric(19,6),  
            INJB09      Numeric(19,6),  
            SUMM09      Numeric(19,6),  
            BT1M09      Numeric(19,6),  
            BT2M09      Numeric(19,6),  
            BT3M09      Numeric(19,6),      
            GABM09      Numeric(19,6),      
            JUMM09      Numeric(19,6),  
            KUKM09      Numeric(19,6),  
            MEDM09      Numeric(19,6),  
            GBHM09      Numeric(19,6),
            JUSM09      Numeric(19,6),
            GBUM09      Numeric(19,6),
            BU3M09      Numeric(19,6),
            BT4M09      Numeric(19,6),
            BT5M09      Numeric(19,6),
            BT6M09      Numeric(19,6),
            JIGM10      nvarchar(6),    
            PAYM10      Numeric(19,6),  
            BNSM10      Numeric(19,6),  
            INJB10      Numeric(19,6),  
            SUMM10      Numeric(19,6),  
            BT1M10      Numeric(19,6),  
            BT2M10      Numeric(19,6),  
            BT3M10      Numeric(19,6),      
            GABM10      Numeric(19,6),      
            JUMM10      Numeric(19,6),  
            KUKM10      Numeric(19,6),  
            MEDM10      Numeric(19,6),  
            GBHM10      Numeric(19,6),
            JUSM10      Numeric(19,6),
            GBUM10      Numeric(19,6),
            BU3M10      Numeric(19,6),
            BT4M10      Numeric(19,6),
            BT5M10      Numeric(19,6),
            BT6M10      Numeric(19,6),
            JIGM11      nvarchar(6),    
            PAYM11      Numeric(19,6),  
            BNSM11      Numeric(19,6),  
            INJB11      Numeric(19,6),  
            SUMM11      Numeric(19,6),  
            BT1M11      Numeric(19,6),  
            BT2M11      Numeric(19,6),  
            BT3M11      Numeric(19,6),      
            GABM11      Numeric(19,6),      
            JUMM11      Numeric(19,6),  
            KUKM11      Numeric(19,6),  
            MEDM11      Numeric(19,6),  
            GBHM11      Numeric(19,6),
            JUSM11      Numeric(19,6),
            GBUM11      Numeric(19,6),
            BU3M11      Numeric(19,6),
            BT4M11      Numeric(19,6),
            BT5M11      Numeric(19,6),
            BT6M11      Numeric(19,6),
            JIGM12      nvarchar(6),    
            PAYM12      Numeric(19,6),  
            BNSM12      Numeric(19,6),  
            INJB12      Numeric(19,6),  
            SUMM12      Numeric(19,6),  
            BT1M12      Numeric(19,6),  
            BT2M12      Numeric(19,6),  
            BT3M12      Numeric(19,6),      
            GABM12      Numeric(19,6),      
            JUMM12      Numeric(19,6),  
            KUKM12      Numeric(19,6),  
            MEDM12      Numeric(19,6),  
            GBHM12      Numeric(19,6),
            JUSM12      Numeric(19,6),
            GBUM12      Numeric(19,6),
            BU3M12      Numeric(19,6),
            BT4M12      Numeric(19,6),
            BT5M12      Numeric(19,6),
            BT6M12      Numeric(19,6),
            JIGM13      nvarchar(6),    
            PAYM13      Numeric(19,6),  
            BNSM13      Numeric(19,6),  
            INJB13      Numeric(19,6),  
            SUMM13      Numeric(19,6),  
            BT1M13      Numeric(19,6),  
            BT2M13      Numeric(19,6),  
            BT3M13      Numeric(19,6),      
            GABM13      Numeric(19,6),      
            JUMM13      Numeric(19,6),  
            KUKM13      Numeric(19,6),  
            MEDM13      Numeric(19,6),  
            GBHM13      Numeric(19,6),
            JUSM13      Numeric(19,6),
            GBUM13      Numeric(19,6),
            BU3M13      Numeric(19,6),
            BT4M13      Numeric(19,6),
            BT5M13      Numeric(19,6),
            BT6M13      Numeric(19,6),
            CLTNAM      nvarchar(50),
            COMPRT      nvarchar(30),
            BUSNUM      nvarchar(15),
            PERNUM      nvarchar(15),
            POSADD      nvarchar(100),
            URIB01      Numeric(19,6),
            URIB02      Numeric(19,6),
            URIB03      Numeric(19,6),
            URIB04      Numeric(19,6),
            URIB05      Numeric(19,6),
            URIB06      Numeric(19,6),
            URIB07      Numeric(19,6),
            URIB08      Numeric(19,6),
            URIB09      Numeric(19,6),
            URIB10      Numeric(19,6),
            URIB11      Numeric(19,6),
            URIB12      Numeric(19,6),
            URIB13      Numeric(19,6),
            BT7M01      Numeric(19,6),
            BT7M02      Numeric(19,6),
            BT7M03      Numeric(19,6),
            BT7M04      Numeric(19,6),
            BT7M05      Numeric(19,6),
            BT7M06      Numeric(19,6),
            BT7M07      Numeric(19,6),
            BT7M08      Numeric(19,6),
            BT7M09      Numeric(19,6),
            BT7M10      Numeric(19,6),
            BT7M11      Numeric(19,6),
            BT7M12      Numeric(19,6),
            BT7M13      Numeric(19,6),
            BT8M01      Numeric(19,6),
            BT8M02      Numeric(19,6),
            BT8M03      Numeric(19,6),
            BT8M04      Numeric(19,6),
            BT8M05      Numeric(19,6),
            BT8M06      Numeric(19,6),
            BT8M07      Numeric(19,6),
            BT8M08      Numeric(19,6),
            BT8M09      Numeric(19,6),
            BT8M10      Numeric(19,6),
            BT8M11      Numeric(19,6),
            BT8M12      Numeric(19,6),
            BT8M13      Numeric(19,6)
            ) 

-- <1.월별 자료 조회 >ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ    
    INSERT INTO [#RPY503]
    SELECT  T1.U_MstCode	AS MSTCOD,                                                           
            T1.U_MstName	AS MSTNAM,                                                           
            T2.U_StartDat	AS STRDAY,
			T2.U_termDate	AS ENDDAY,                                                            
            T2.U_govID		AS PERNBR,                                                           
            (1+ (case when T2.U_BAEWOO = 'Y' THEN 1 ELSE 0 END) + T2.U_BUYNSU) AS BUYNSU,               

            T0.JIGM01, T0.PAYM01, T0.BNSM01, T0.INJB01, T0.SUMM01, T0.BT1M01, T0.BT2M01, T0.BT3M01, T0.GABM01, T0.JUMM01, T0.KUKM01, T0.MEDM01, T0.GBHM01, T0.JUSM01, T0.GBUM01, T0.BU3M01, T0.BT4M01, T0.BT5M01, T0.BT6M01,        
            T0.JIGM02, T0.PAYM02, T0.BNSM02, T0.INJB02, T0.SUMM02, T0.BT1M02, T0.BT2M02, T0.BT3M02, T0.GABM02, T0.JUMM02, T0.KUKM02, T0.MEDM02, T0.GBHM02, T0.JUSM02, T0.GBUM02, T0.BU3M02, T0.BT4M02, T0.BT5M02, T0.BT6M02,
            T0.JIGM03, T0.PAYM03, T0.BNSM03, T0.INJB03, T0.SUMM03, T0.BT1M03, T0.BT2M03, T0.BT3M03, T0.GABM03, T0.JUMM03, T0.KUKM03, T0.MEDM03, T0.GBHM03, T0.JUSM03, T0.GBUM03, T0.BU3M03, T0.BT4M03, T0.BT5M03, T0.BT6M03,
            T0.JIGM04, T0.PAYM04, T0.BNSM04, T0.INJB04, T0.SUMM04, T0.BT1M04, T0.BT2M04, T0.BT3M04, T0.GABM04, T0.JUMM04, T0.KUKM04, T0.MEDM04, T0.GBHM04, T0.JUSM04, T0.GBUM04, T0.BU3M04, T0.BT4M04, T0.BT5M04, T0.BT6M04,
            T0.JIGM05, T0.PAYM05, T0.BNSM05, T0.INJB05, T0.SUMM05, T0.BT1M05, T0.BT2M05, T0.BT3M05, T0.GABM05, T0.JUMM05, T0.KUKM05, T0.MEDM05, T0.GBHM05, T0.JUSM05, T0.GBUM05, T0.BU3M05, T0.BT4M05, T0.BT5M05, T0.BT6M05,
            T0.JIGM06, T0.PAYM06, T0.BNSM06, T0.INJB06, T0.SUMM06, T0.BT1M06, T0.BT2M06, T0.BT3M06, T0.GABM06, T0.JUMM06, T0.KUKM06, T0.MEDM06, T0.GBHM06, T0.JUSM06, T0.GBUM06, T0.BU3M06, T0.BT4M06, T0.BT5M06, T0.BT6M06,
            T0.JIGM07, T0.PAYM07, T0.BNSM07, T0.INJB07, T0.SUMM07, T0.BT1M07, T0.BT2M07, T0.BT3M07, T0.GABM07, T0.JUMM07, T0.KUKM07, T0.MEDM07, T0.GBHM07, T0.JUSM07, T0.GBUM07, T0.BU3M07, T0.BT4M07, T0.BT5M07, T0.BT6M07,
            T0.JIGM08, T0.PAYM08, T0.BNSM08, T0.INJB08, T0.SUMM08, T0.BT1M08, T0.BT2M08, T0.BT3M08, T0.GABM08, T0.JUMM08, T0.KUKM08, T0.MEDM08, T0.GBHM08, T0.JUSM08, T0.GBUM08, T0.BU3M08, T0.BT4M08, T0.BT5M08, T0.BT6M08,
            T0.JIGM09, T0.PAYM09, T0.BNSM09, T0.INJB09, T0.SUMM09, T0.BT1M09, T0.BT2M09, T0.BT3M09, T0.GABM09, T0.JUMM09, T0.KUKM09, T0.MEDM09, T0.GBHM09, T0.JUSM09, T0.GBUM09, T0.BU3M09, T0.BT4M09, T0.BT5M09, T0.BT6M09,
            T0.JIGM10, T0.PAYM10, T0.BNSM10, T0.INJB10, T0.SUMM10, T0.BT1M10, T0.BT2M10, T0.BT3M10, T0.GABM10, T0.JUMM10, T0.KUKM10, T0.MEDM10, T0.GBHM10, T0.JUSM10, T0.GBUM10, T0.BU3M10, T0.BT4M10, T0.BT5M10, T0.BT6M10,
            T0.JIGM11, T0.PAYM11, T0.BNSM11, T0.INJB11, T0.SUMM11, T0.BT1M11, T0.BT2M11, T0.BT3M11, T0.GABM11, T0.JUMM11, T0.KUKM11, T0.MEDM11, T0.GBHM11, T0.JUSM11, T0.GBUM11, T0.BU3M11, T0.BT4M11, T0.BT5M11, T0.BT6M11,
            T0.JIGM12, T0.PAYM12, T0.BNSM12, T0.INJB12, T0.SUMM12, T0.BT1M12, T0.BT2M12, T0.BT3M12, T0.GABM12, T0.JUMM12, T0.KUKM12, T0.MEDM12, T0.GBHM12, T0.JUSM12, T0.GBUM12, T0.BU3M12, T0.BT4M12, T0.BT5M12, T0.BT6M12,
            T0.JIGM13, T0.PAYM13, T0.BNSM13, T0.INJB13, T0.SUMM13, T0.BT1M13, T0.BT2M13, T0.BT3M13, T0.GABM13, T0.JUMM13, T0.KUKM13, T0.MEDM13, T0.GBHM13, T0.JUSM13, T0.GBUM13, T0.BU3M13, T0.BT4M13, T0.BT5M13, T0.BT6M13,
            ISNULL(T4.U_CLTName, '') AS CLTNAM,
            ISNULL(T4.U_ComPrt, '') AS COMPRT,
            ISNULL(T4.U_BusNum, '') AS BUSNUM,
            ISNULL(T4.U_PerNum, '') AS PERNUM,
            ISNULL(T4.U_PosAdd, '') AS POSADD,
            T0.URIB01, T0.URIB02, T0.URIB03, T0.URIB04, T0.URIB05, T0.URIB06, T0.URIB07, T0.URIB08, T0.URIB09, T0.URIB10, T0.URIB11, T0.URIB12, T0.URIB13, 
            T0.BT7M01, T0.BT7M02, T0.BT7M03, T0.BT7M04, T0.BT7M05, T0.BT7M06, T0.BT7M07, T0.BT7M08, T0.BT7M09, T0.BT7M10, T0.BT7M11, T0.BT7M12, T0.BT7M13,
            T0.BT8M01, T0.BT8M02, T0.BT8M03, T0.BT8M04, T0.BT8M05, T0.BT8M06, T0.BT8M07, T0.BT8M08, T0.BT8M09, T0.BT8M10, T0.BT8M11, T0.BT8M12, T0.BT8M13
    --INTO [#RPY503]
    FROM 
        (
            SELECT  T01.*,  T02.*,  T03.*,  T04.*,  T05.*,  T06.*, 
                    T07.*,  T08.*,  T09.*,  T10.*,  T11.*,  T12.*,
                    T13.*
            FROM    (
                SELECT DocEntry,    
                       JIGM01 =  MAX(SUBSTRING(U_JIGDATE, 1, 6)),
                       PAYM01 =  SUM(U_GwaPay), 
                       BNSM01 =  SUM(U_GwaBns), 
                       INJB01 =  SUM(U_InJBns),
                       SUMM01 =  SUM(U_GwaSee),
                       BT3M01 =  SUM(ISNULL(U_BiGwa03,0) + ISNULL(U_BIGM01,0)+ ISNULL(U_BIGM02,0)+ ISNULL(U_BIGM03,0)),
                       BT1M01 =  SUM(ISNULL(U_BiGwa01,0) + ISNULL(U_BIGO01,0)),
                       BT2M01 =  SUM(U_BiGwa02 + ISNULL(U_BiGwa07,0)),
                       GABM01 =  SUM(U_GabGun),
                       JUMM01 =  SUM(U_JuMin) ,
                       KUKM01 =  SUM(U_KukAmt),
                       MEDM01 =  SUM(U_MedAmt + ISNULL(U_NGYAMT,0)),
                       GBHM01 =  SUM(U_GBHAMT),
                       JUSM01 =  SUM(U_JUSBNS),
                       GBUM01 =  SUM(ISNULL(U_GbuAmt,0)) ,
                       BU3M01 =  SUM(ISNULL(U_BiGwu03,0) + ISNULL(U_BIGX01,0)),
                       BT4M01 =  SUM(ISNULL(U_BiGwa04,0) + ISNULL(U_BIGG01,0) + ISNULL(U_BIGH01,0) + ISNULL(U_BIGH11,0) 
                                    + ISNULL(U_BIGH12,0) + ISNULL(U_BIGH13,0) + ISNULL(U_BIGI01,0) + ISNULL(U_BIGK01,0) 
                                    + ISNULL(U_BIGS01,0) + ISNULL(U_BIGT01,0) + ISNULL(U_BIGY01,0) + ISNULL(U_BIGY02,0) 
                                    + ISNULL(U_BIGY03,0) + ISNULL(U_BIGY20,0) + ISNULL(U_BIGY21,0) + ISNULL(U_BIGZ01,0)),
                       BT5M01 =  SUM(ISNULL(U_BiGwa05,0) + ISNULL(U_BIGH06,0) + ISNULL(U_BIGH07,0) + ISNULL(U_BIGH08,0) + ISNULL(U_BIGH09,0) + ISNULL(U_BIGH10,0)),
                       BT6M01 =  SUM(ISNULL(U_BiGwa06,0) + ISNULL(U_BIGQ01,0)),
                       URIB01 =  SUM(ISNULL(U_URIBNS,0)),
                       BT7M01 =  SUM(ISNULL(U_BiGwa02,0)),
                       BT8M01 =  SUM(ISNULL(U_BiGwa07,0))
                FROM   [@ZPY343L]
                WHERE  U_LineNum = '01'
                GROUP  BY DocEntry
                ) T01,
                (
                SELECT DocM02 =  DocEntry,
                       JIGM02 =  MAX(SUBSTRING(U_JIGDATE, 1, 6))     ,
                       PAYM02 =  SUM(U_GwaPay),
                       BNSM02 =  SUM(U_GwaBns),
                       INJB02 =  SUM(U_InJBns),
                       SUMM02 =  SUM(U_GwaSee),
                       BT3M02 =  SUM(ISNULL(U_BiGwa03,0) + ISNULL(U_BIGM01,0)+ ISNULL(U_BIGM02,0)+ ISNULL(U_BIGM03,0)),
                       BT1M02 =  SUM(ISNULL(U_BiGwa01,0) + ISNULL(U_BIGO01,0)),
                       BT2M02 =  SUM(U_BiGwa02 + ISNULL(U_BiGwa07,0)),
                       GABM02 =  SUM(U_GabGun),
                       JUMM02 =  SUM(U_JuMin),
                       KUKM02 =  SUM(U_KukAmt),
                       MEDM02 =  SUM(U_MedAmt + ISNULL(U_NGYAMT,0)),
                       GBHM02 =  SUM(U_GBHAMT),
                       JUSM02 =  SUM(U_JUSBNS),
                       GBUM02 =  SUM(ISNULL(U_GbuAmt,0)) ,
                       BU3M02 =  SUM(ISNULL(U_BiGwu03,0) + ISNULL(U_BIGX01,0)),
                       BT4M02 =  SUM(ISNULL(U_BiGwa04,0) + ISNULL(U_BIGG01,0) + ISNULL(U_BIGH01,0) + ISNULL(U_BIGH11,0) 
                                    + ISNULL(U_BIGH12,0) + ISNULL(U_BIGH13,0) + ISNULL(U_BIGI01,0) + ISNULL(U_BIGK01,0) 
                                    + ISNULL(U_BIGS01,0) + ISNULL(U_BIGT01,0) + ISNULL(U_BIGY01,0) + ISNULL(U_BIGY02,0) 
                                    + ISNULL(U_BIGY03,0) + ISNULL(U_BIGY20,0) + ISNULL(U_BIGY21,0) + ISNULL(U_BIGZ01,0)),
                       BT5M02 =  SUM(ISNULL(U_BiGwa05,0) + ISNULL(U_BIGH06,0) + ISNULL(U_BIGH07,0) + ISNULL(U_BIGH08,0) + ISNULL(U_BIGH09,0) + ISNULL(U_BIGH10,0))          ,
                       BT6M02 =  SUM(ISNULL(U_BiGwa06,0) + ISNULL(U_BIGQ01,0)),
                       URIB02 =  SUM(ISNULL(U_URIBNS,0)),
                       BT7M02 =  SUM(ISNULL(U_BiGwa02,0)),
                       BT8M02 =  SUM(ISNULL(U_BiGwa07,0))
                FROM   [@ZPY343L]
                WHERE  U_LineNum = '02'
                GROUP  BY DocEntry
                ) T02,
                (
                SELECT DocM03 =  DocEntry,
                       JIGM03 =  MAX(SUBSTRING(U_JIGDATE, 1, 6)),
                       PAYM03 =  SUM(U_GwaPay),
                       BNSM03 =  SUM(U_GwaBns),
                       INJB03 =  SUM(U_InJBns),
                       SUMM03 =  SUM(U_GwaSee),
                       BT3M03 =  SUM(ISNULL(U_BiGwa03,0) + ISNULL(U_BIGM01,0)+ ISNULL(U_BIGM02,0)+ ISNULL(U_BIGM03,0)),
                       BT1M03 =  SUM(ISNULL(U_BiGwa01,0) + ISNULL(U_BIGO01,0)),
                       BT2M03 =  SUM(U_BiGwa02 + ISNULL(U_BiGwa07,0)),
                       GABM03 =  SUM(U_GabGun),
                       JUMM03 =  SUM(U_JuMin),
                       KUKM03 =  SUM(U_KukAmt),
                       MEDM03 =  SUM(U_MedAmt + ISNULL(U_NGYAMT,0)),
                       GBHM03 =  SUM(U_GBHAMT),
                       JUSM03 =  SUM(U_JUSBNS),
                       GBUM03 =  SUM(ISNULL(U_GbuAmt,0)) ,
                       BU3M03 =  SUM(ISNULL(U_BiGwu03,0) + ISNULL(U_BIGX01,0)),
                       BT4M03 =  SUM(ISNULL(U_BiGwa04,0) + ISNULL(U_BIGG01,0) + ISNULL(U_BIGH01,0) + ISNULL(U_BIGH11,0) 
                                    + ISNULL(U_BIGH12,0) + ISNULL(U_BIGH13,0) + ISNULL(U_BIGI01,0) + ISNULL(U_BIGK01,0) 
                                    + ISNULL(U_BIGS01,0) + ISNULL(U_BIGT01,0) + ISNULL(U_BIGY01,0) + ISNULL(U_BIGY02,0) 
                                    + ISNULL(U_BIGY03,0) + ISNULL(U_BIGY20,0) + ISNULL(U_BIGY21,0) + ISNULL(U_BIGZ01,0)),
                       BT5M03 =  SUM(ISNULL(U_BiGwa05,0) + ISNULL(U_BIGH06,0) + ISNULL(U_BIGH07,0) + ISNULL(U_BIGH08,0) + ISNULL(U_BIGH09,0) + ISNULL(U_BIGH10,0))          ,
                       BT6M03 =  SUM(ISNULL(U_BiGwa06,0) + ISNULL(U_BIGQ01,0)),
                       URIB03 =  SUM(ISNULL(U_URIBNS,0)),
                       BT7M03 =  SUM(ISNULL(U_BiGwa02,0)),
                       BT8M03 =  SUM(ISNULL(U_BiGwa07,0))
                FROM   [@ZPY343L]
                WHERE  U_LineNum = '03'
                GROUP  BY DocEntry
                ) T03,
                (
                SELECT DocM04 =  DocEntry,
                       JIGM04 =  MAX(SUBSTRING(U_JIGDATE, 1, 6)),
                       PAYM04 =  SUM(U_GwaPay),
                       BNSM04 =  SUM(U_GwaBns),
                       INJB04 =  SUM(U_InJBns),
                       SUMM04 =  SUM(U_GwaSee),
                       BT3M04 =  SUM(ISNULL(U_BiGwa03,0) + ISNULL(U_BIGM01,0)+ ISNULL(U_BIGM02,0)+ ISNULL(U_BIGM03,0)),
                       BT1M04 =  SUM(ISNULL(U_BiGwa01,0) + ISNULL(U_BIGO01,0)),
                       BT2M04 =  SUM(U_BiGwa02 + ISNULL(U_BiGwa07,0)),
                       GABM04 =  SUM(U_GabGun),
                       JUMM04 =  SUM(U_JuMin),
                       KUKM04 =  SUM(U_KukAmt),
                       MEDM04 =  SUM(U_MedAmt + ISNULL(U_NGYAMT,0)),
                       GBHM04 =  SUM(U_GBHAMT),
                       JUSM04 =  SUM(U_JUSBNS),
                       GBUM04 =  SUM(ISNULL(U_GbuAmt,0)) ,
                       BU3M04 =  SUM(ISNULL(U_BiGwu03,0) + ISNULL(U_BIGX01,0)),
                       BT4M04 =  SUM(ISNULL(U_BiGwa04,0) + ISNULL(U_BIGG01,0) + ISNULL(U_BIGH01,0) + ISNULL(U_BIGH11,0) 
                                    + ISNULL(U_BIGH12,0) + ISNULL(U_BIGH13,0) + ISNULL(U_BIGI01,0) + ISNULL(U_BIGK01,0) 
                                    + ISNULL(U_BIGS01,0) + ISNULL(U_BIGT01,0) + ISNULL(U_BIGY01,0) + ISNULL(U_BIGY02,0) 
                                    + ISNULL(U_BIGY03,0) + ISNULL(U_BIGY20,0) + ISNULL(U_BIGY21,0) + ISNULL(U_BIGZ01,0)),
                       BT5M04 =  SUM(ISNULL(U_BiGwa05,0) + ISNULL(U_BIGH06,0) + ISNULL(U_BIGH07,0) + ISNULL(U_BIGH08,0) + ISNULL(U_BIGH09,0) + ISNULL(U_BIGH10,0))          ,
                       BT6M04 =  SUM(ISNULL(U_BiGwa06,0) + ISNULL(U_BIGQ01,0)),
                       URIB04 =  SUM(ISNULL(U_URIBNS,0)),
                       BT7M04 =  SUM(ISNULL(U_BiGwa02,0)),
                       BT8M04 =  SUM(ISNULL(U_BiGwa07,0))          
                FROM   [@ZPY343L]
                WHERE  U_LineNum = '04'
                GROUP  BY DocEntry
                ) T04,
                (
                SELECT DocM05 =  DocEntry,
                       JIGM05 =  MAX(SUBSTRING(U_JIGDATE, 1, 6)),
                       PAYM05 =  SUM(U_GwaPay),
                       BNSM05 =  SUM(U_GwaBns),
                       INJB05 =  SUM(U_InJBns),
                       SUMM05 =  SUM(U_GwaSee),
                       BT3M05 =  SUM(ISNULL(U_BiGwa03,0) + ISNULL(U_BIGM01,0)+ ISNULL(U_BIGM02,0)+ ISNULL(U_BIGM03,0)),
                       BT1M05 =  SUM(ISNULL(U_BiGwa01,0) + ISNULL(U_BIGO01,0)),
                       BT2M05 =  SUM(U_BiGwa02 + ISNULL(U_BiGwa07,0)),
                       GABM05 =  SUM(U_GabGun),
                       JUMM05 =  SUM(U_JuMin),
                       KUKM05 =  SUM(U_KukAmt),
                       MEDM05 =  SUM(U_MedAmt + ISNULL(U_NGYAMT,0)),
                       GBHM05 =  SUM(U_GBHAMT),
                       JUSM05 =  SUM(U_JUSBNS),
                       GBUM05 =  SUM(ISNULL(U_GbuAmt,0)) ,
                       BU3M05 =  SUM(ISNULL(U_BiGwu03,0) + ISNULL(U_BIGX01,0)),
                       BT4M05 =  SUM(ISNULL(U_BiGwa04,0) + ISNULL(U_BIGG01,0) + ISNULL(U_BIGH01,0) + ISNULL(U_BIGH11,0) 
                                    + ISNULL(U_BIGH12,0) + ISNULL(U_BIGH13,0) + ISNULL(U_BIGI01,0) + ISNULL(U_BIGK01,0) 
                                    + ISNULL(U_BIGS01,0) + ISNULL(U_BIGT01,0) + ISNULL(U_BIGY01,0) + ISNULL(U_BIGY02,0) 
                                    + ISNULL(U_BIGY03,0) + ISNULL(U_BIGY20,0) + ISNULL(U_BIGY21,0) + ISNULL(U_BIGZ01,0)),
                       BT5M05 =  SUM(ISNULL(U_BiGwa05,0) + ISNULL(U_BIGH06,0) + ISNULL(U_BIGH07,0) + ISNULL(U_BIGH08,0) + ISNULL(U_BIGH09,0) + ISNULL(U_BIGH10,0))          ,
                       BT6M05 =  SUM(ISNULL(U_BiGwa06,0) + ISNULL(U_BIGQ01,0)),
                       URIB05 =  SUM(ISNULL(U_URIBNS,0)),
                       BT7M05 =  SUM(ISNULL(U_BiGwa02,0)),
                       BT8M05 =  SUM(ISNULL(U_BiGwa07,0))          
                FROM   [@ZPY343L]
                WHERE  U_LineNum = '05'
                GROUP  BY DocEntry
                ) T05,
                (
                SELECT DocM06 =  DocEntry,
                       JIGM06 =  MAX(SUBSTRING(U_JIGDATE, 1, 6)),
                       PAYM06 =  SUM(U_GwaPay),
                       BNSM06 =  SUM(U_GwaBns),
                       INJB06 =  SUM(U_InJBns),
                       SUMM06 =  SUM(U_GwaSee),
                       BT3M06 =  SUM(ISNULL(U_BiGwa03,0) + ISNULL(U_BIGM01,0)+ ISNULL(U_BIGM02,0)+ ISNULL(U_BIGM03,0)),
                       BT1M06 =  SUM(ISNULL(U_BiGwa01,0) + ISNULL(U_BIGO01,0)),
                       BT2M06 =  SUM(U_BiGwa02 + ISNULL(U_BiGwa07,0)),
                       GABM06 =  SUM(U_GabGun),
                       JUMM06 =  SUM(U_JuMin),
                       KUKM06 =  SUM(U_KukAmt),
                       MEDM06 =  SUM(U_MedAmt + ISNULL(U_NGYAMT,0)),
                       GBHM06 =  SUM(U_GBHAMT),
                       JUSM06 =  SUM(U_JUSBNS),
                       GBUM06 =  SUM(ISNULL(U_GbuAmt,0)) ,
                       BU3M06 =  SUM(ISNULL(U_BiGwu03,0) + ISNULL(U_BIGX01,0)),
                       BT4M06 =  SUM(ISNULL(U_BiGwa04,0) + ISNULL(U_BIGG01,0) + ISNULL(U_BIGH01,0) + ISNULL(U_BIGH11,0) 
                                    + ISNULL(U_BIGH12,0) + ISNULL(U_BIGH13,0) + ISNULL(U_BIGI01,0) + ISNULL(U_BIGK01,0) 
                                    + ISNULL(U_BIGS01,0) + ISNULL(U_BIGT01,0) + ISNULL(U_BIGY01,0) + ISNULL(U_BIGY02,0) 
                                    + ISNULL(U_BIGY03,0) + ISNULL(U_BIGY20,0) + ISNULL(U_BIGY21,0) + ISNULL(U_BIGZ01,0)),
                       BT5M06 =  SUM(ISNULL(U_BiGwa05,0) + ISNULL(U_BIGH06,0) + ISNULL(U_BIGH07,0) + ISNULL(U_BIGH08,0) + ISNULL(U_BIGH09,0) + ISNULL(U_BIGH10,0))          ,
                       BT6M06 =  SUM(ISNULL(U_BiGwa06,0) + ISNULL(U_BIGQ01,0)),
                       URIB06 =  SUM(ISNULL(U_URIBNS,0)),
                       BT7M06 =  SUM(ISNULL(U_BiGwa02,0)),
                       BT8M06 =  SUM(ISNULL(U_BiGwa07,0))  
                FROM   [@ZPY343L]
                WHERE  U_LineNum = '06'
                GROUP  BY DocEntry
                ) T06,
                (
                SELECT DocM07 =  DocEntry,
                       JIGM07 =  MAX(SUBSTRING(U_JIGDATE, 1, 6)),
                       PAYM07 =  SUM(U_GwaPay),
                       BNSM07 =  SUM(U_GwaBns),
                       INJB07 =  SUM(U_InJBns),
                       SUMM07 =  SUM(U_GwaSee),
                       BT3M07 =  SUM(ISNULL(U_BiGwa03,0) + ISNULL(U_BIGM01,0)+ ISNULL(U_BIGM02,0)+ ISNULL(U_BIGM03,0)),
                       BT1M07 =  SUM(ISNULL(U_BiGwa01,0) + ISNULL(U_BIGO01,0)),
                       BT2M07 =  SUM(U_BiGwa02 + ISNULL(U_BiGwa07,0)),
                       GABM07 =  SUM(U_GabGun),
                       JUMM07 =  SUM(U_JuMin),
                       KUKM07 =  SUM(U_KukAmt),
                       MEDM07 =  SUM(U_MedAmt + ISNULL(U_NGYAMT,0)),
                       GBHM07 =  SUM(U_GBHAMT),
                       JUSM07 =  SUM(U_JUSBNS),
                       GBUM07 =  SUM(ISNULL(U_GbuAmt,0)) ,
                       BU3M07 =  SUM(ISNULL(U_BiGwu03,0) + ISNULL(U_BIGX01,0)),
                       BT4M07 =  SUM(ISNULL(U_BiGwa04,0) + ISNULL(U_BIGG01,0) + ISNULL(U_BIGH01,0) + ISNULL(U_BIGH11,0) 
                                    + ISNULL(U_BIGH12,0) + ISNULL(U_BIGH13,0) + ISNULL(U_BIGI01,0) + ISNULL(U_BIGK01,0) 
                                    + ISNULL(U_BIGS01,0) + ISNULL(U_BIGT01,0) + ISNULL(U_BIGY01,0) + ISNULL(U_BIGY02,0) 
                                    + ISNULL(U_BIGY03,0) + ISNULL(U_BIGY20,0) + ISNULL(U_BIGY21,0) + ISNULL(U_BIGZ01,0)),
                       BT5M07 =  SUM(ISNULL(U_BiGwa05,0) + ISNULL(U_BIGH06,0) + ISNULL(U_BIGH07,0) + ISNULL(U_BIGH08,0) + ISNULL(U_BIGH09,0) + ISNULL(U_BIGH10,0))          ,
                       BT6M07 =  SUM(ISNULL(U_BiGwa06,0) + ISNULL(U_BIGQ01,0)),
                       URIB07 =  SUM(ISNULL(U_URIBNS,0)),
                       BT7M07 =  SUM(ISNULL(U_BiGwa02,0)),
                       BT8M07 =  SUM(ISNULL(U_BiGwa07,0))          
                FROM   [@ZPY343L]
                WHERE  U_LineNum = '07'
                GROUP  BY DocEntry
                ) T07,
                (
                SELECT DocM08 =  DocEntry,
                       JIGM08 =  MAX(SUBSTRING(U_JIGDATE, 1, 6)),
                       PAYM08 =  SUM(U_GwaPay) ,
                       BNSM08 =  SUM(U_GwaBns) ,
                       INJB08 =  SUM(U_InJBns) ,
                       SUMM08 =  SUM(U_GwaSee) ,
                       BT3M08 =  SUM(ISNULL(U_BiGwa03,0) + ISNULL(U_BIGM01,0)+ ISNULL(U_BIGM02,0)+ ISNULL(U_BIGM03,0)),
                       BT1M08 =  SUM(ISNULL(U_BiGwa01,0) + ISNULL(U_BIGO01,0)),
                       BT2M08 =  SUM(U_BiGwa02 + ISNULL(U_BiGwa07,0)),
                       GABM08 =  SUM(U_GabGun),
                       JUMM08 =  SUM(U_JuMin),
                       KUKM08 =  SUM(U_KukAmt),
                       MEDM08 =  SUM(U_MedAmt + ISNULL(U_NGYAMT,0)),
                       GBHM08 =  SUM(U_GBHAMT),
                       JUSM08 =  SUM(U_JUSBNS),
                       GBUM08 =  SUM(ISNULL(U_GbuAmt,0)) ,
                       BU3M08 =  SUM(ISNULL(U_BiGwu03,0) + ISNULL(U_BIGX01,0)),
                       BT4M08 =  SUM(ISNULL(U_BiGwa04,0) + ISNULL(U_BIGG01,0) + ISNULL(U_BIGH01,0) + ISNULL(U_BIGH11,0) 
                                    + ISNULL(U_BIGH12,0) + ISNULL(U_BIGH13,0) + ISNULL(U_BIGI01,0) + ISNULL(U_BIGK01,0) 
                                    + ISNULL(U_BIGS01,0) + ISNULL(U_BIGT01,0) + ISNULL(U_BIGY01,0) + ISNULL(U_BIGY02,0) 
                                    + ISNULL(U_BIGY03,0) + ISNULL(U_BIGY20,0) + ISNULL(U_BIGY21,0) + ISNULL(U_BIGZ01,0)),
                       BT5M08 =  SUM(ISNULL(U_BiGwa05,0) + ISNULL(U_BIGH06,0) + ISNULL(U_BIGH07,0) + ISNULL(U_BIGH08,0) + ISNULL(U_BIGH09,0) + ISNULL(U_BIGH10,0))          ,
                       BT6M08 =  SUM(ISNULL(U_BiGwa06,0) + ISNULL(U_BIGQ01,0)),
                       URIB08 =  SUM(ISNULL(U_URIBNS,0)),
                       BT7M08 =  SUM(ISNULL(U_BiGwa02,0)),
                       BT8M08 =  SUM(ISNULL(U_BiGwa07,0))          
                FROM   [@ZPY343L]
                WHERE  U_LineNum = '08'
                GROUP  BY DocEntry
                ) T08,
    
                (
                SELECT DocM09 =  DocEntry,
                       JIGM09 =  MAX(SUBSTRING(U_JIGDATE, 1, 6)),
                       PAYM09 =  SUM(U_GwaPay) ,
                       BNSM09 =  SUM(U_GwaBns) ,
                       INJB09 =  SUM(U_InJBns) ,
                       SUMM09 =  SUM(U_GwaSee) ,
                       BT3M09 =  SUM(ISNULL(U_BiGwa03,0) + ISNULL(U_BIGM01,0)+ ISNULL(U_BIGM02,0)+ ISNULL(U_BIGM03,0)),
                       BT1M09 =  SUM(ISNULL(U_BiGwa01,0) + ISNULL(U_BIGO01,0)),
                       BT2M09 =  SUM(U_BiGwa02 + ISNULL(U_BiGwa07,0)),
                       GABM09 =  SUM(U_GabGun),
                       JUMM09 =  SUM(U_JuMin),
                       KUKM09 =  SUM(U_KukAmt),
                       MEDM09 =  SUM(U_MedAmt + ISNULL(U_NGYAMT,0)),
                       GBHM09 =  SUM(U_GBHAMT),
                       JUSM09 =  SUM(U_JUSBNS),
                       GBUM09 =  SUM(ISNULL(U_GbuAmt,0)) ,
                       BU3M09 =  SUM(ISNULL(U_BiGwu03,0) + ISNULL(U_BIGX01,0)),
                       BT4M09 =  SUM(ISNULL(U_BiGwa04,0) + ISNULL(U_BIGG01,0) + ISNULL(U_BIGH01,0) + ISNULL(U_BIGH11,0) 
                                    + ISNULL(U_BIGH12,0) + ISNULL(U_BIGH13,0) + ISNULL(U_BIGI01,0) + ISNULL(U_BIGK01,0) 
                                    + ISNULL(U_BIGS01,0) + ISNULL(U_BIGT01,0) + ISNULL(U_BIGY01,0) + ISNULL(U_BIGY02,0) 
                                    + ISNULL(U_BIGY03,0) + ISNULL(U_BIGY20,0) + ISNULL(U_BIGY21,0) + ISNULL(U_BIGZ01,0)),
                       BT5M09 =  SUM(ISNULL(U_BiGwa05,0) + ISNULL(U_BIGH06,0) + ISNULL(U_BIGH07,0) + ISNULL(U_BIGH08,0) + ISNULL(U_BIGH09,0) + ISNULL(U_BIGH10,0)),
                       BT6M09 =  SUM(ISNULL(U_BiGwa06,0) + ISNULL(U_BIGQ01,0)),
                       URIB09 =  SUM(ISNULL(U_URIBNS,0)),
                       BT7M09 =  SUM(ISNULL(U_BiGwa02,0)),
                       BT8M09 =  SUM(ISNULL(U_BiGwa07,0)) 
                FROM   [@ZPY343L]
                WHERE  U_LineNum = '09'
                GROUP  BY DocEntry
                ) T09,
                (
                SELECT DocM10 =  DocEntry,
                       JIGM10 =  MAX(SUBSTRING(U_JIGDATE, 1, 6)),
                       PAYM10 =  SUM(U_GwaPay) ,
                       BNSM10 =  SUM(U_GwaBns) ,
                       INJB10 =  SUM(U_InJBns) ,
                       SUMM10 =  SUM(U_GwaSee) ,
                       BT3M10 =  SUM(ISNULL(U_BiGwa03,0) + ISNULL(U_BIGM01,0)+ ISNULL(U_BIGM02,0)+ ISNULL(U_BIGM03,0)),
                       BT1M10 =  SUM(ISNULL(U_BiGwa01,0) + ISNULL(U_BIGO01,0)),
                       BT2M10 =  SUM(U_BiGwa02 + ISNULL(U_BiGwa07,0)),
                       GABM10 =  SUM(U_GabGun),
                       JUMM10 =  SUM(U_JuMin),
                       KUKM10 =  SUM(U_KukAmt),
                       MEDM10 =  SUM(U_MedAmt + ISNULL(U_NGYAMT,0)),
                       GBHM10 =  SUM(U_GBHAMT),
                       JUSM10 =  SUM(U_JUSBNS),
                       GBUM10 =  SUM(ISNULL(U_GbuAmt,0)) ,
                       BU3M10 =  SUM(ISNULL(U_BiGwu03,0) + ISNULL(U_BIGX01,0)),
                       BT4M10 =  SUM(ISNULL(U_BiGwa04,0) + ISNULL(U_BIGG01,0) + ISNULL(U_BIGH01,0) + ISNULL(U_BIGH11,0) 
                                    + ISNULL(U_BIGH12,0) + ISNULL(U_BIGH13,0) + ISNULL(U_BIGI01,0) + ISNULL(U_BIGK01,0) 
                                    + ISNULL(U_BIGS01,0) + ISNULL(U_BIGT01,0) + ISNULL(U_BIGY01,0) + ISNULL(U_BIGY02,0) 
                                    + ISNULL(U_BIGY03,0) + ISNULL(U_BIGY20,0) + ISNULL(U_BIGY21,0) + ISNULL(U_BIGZ01,0)),
                       BT5M10 =  SUM(ISNULL(U_BiGwa05,0) + ISNULL(U_BIGH06,0) + ISNULL(U_BIGH07,0) + ISNULL(U_BIGH08,0) + ISNULL(U_BIGH09,0) + ISNULL(U_BIGH10,0)),
                       BT6M10 =  SUM(ISNULL(U_BiGwa06,0) + ISNULL(U_BIGQ01,0)),
                       URIB10 =  SUM(ISNULL(U_URIBNS,0)),
                       BT7M10 =  SUM(ISNULL(U_BiGwa02,0)),
                       BT8M10 =  SUM(ISNULL(U_BiGwa07,0))
                FROM   [@ZPY343L]
                WHERE  U_LineNum = '10'
                GROUP  BY DocEntry
                ) T10,
                (
                SELECT DocM11 =  DocEntry,
                       JIGM11 =  MAX(SUBSTRING(U_JIGDATE, 1, 6)),
                       PAYM11 =  SUM(U_GwaPay) ,
                       BNSM11 =  SUM(U_GwaBns) ,
                       INJB11 =  SUM(U_InJBns) ,
                       SUMM11 =  SUM(U_GwaSee) ,
                       BT3M11 =  SUM(ISNULL(U_BiGwa03,0) + ISNULL(U_BIGM01,0)+ ISNULL(U_BIGM02,0)+ ISNULL(U_BIGM03,0)),
                       BT1M11 =  SUM(ISNULL(U_BiGwa01,0) + ISNULL(U_BIGO01,0)),
                       BT2M11 =  SUM(U_BiGwa02 + ISNULL(U_BiGwa07,0)),
                       GABM11 =  SUM(U_GabGun),
                       JUMM11 =  SUM(U_JuMin) ,
                       KUKM11 =  SUM(U_KukAmt),
                       MEDM11 =  SUM(U_MedAmt + ISNULL(U_NGYAMT,0)),
                       GBHM11 =  SUM(U_GBHAMT),
                       JUSM11 =  SUM(U_JUSBNS),
                       GBUM11 =  SUM(ISNULL(U_GbuAmt,0)) ,
                       BU3M11 =  SUM(ISNULL(U_BiGwu03,0) + ISNULL(U_BIGX01,0)),
                       BT4M11 =  SUM(ISNULL(U_BiGwa04,0) + ISNULL(U_BIGG01,0) + ISNULL(U_BIGH01,0) + ISNULL(U_BIGH11,0) 
                                    + ISNULL(U_BIGH12,0) + ISNULL(U_BIGH13,0) + ISNULL(U_BIGI01,0) + ISNULL(U_BIGK01,0) 
                                    + ISNULL(U_BIGS01,0) + ISNULL(U_BIGT01,0) + ISNULL(U_BIGY01,0) + ISNULL(U_BIGY02,0) 
                                    + ISNULL(U_BIGY03,0) + ISNULL(U_BIGY20,0) + ISNULL(U_BIGY21,0) + ISNULL(U_BIGZ01,0)),
                       BT5M11 =  SUM(ISNULL(U_BiGwa05,0) + ISNULL(U_BIGH06,0) + ISNULL(U_BIGH07,0) + ISNULL(U_BIGH08,0) + ISNULL(U_BIGH09,0) + ISNULL(U_BIGH10,0)),
                       BT6M11 =  SUM(ISNULL(U_BiGwa06,0) + ISNULL(U_BIGQ01,0)),
                       URIB11 =  SUM(ISNULL(U_URIBNS,0)),
                       BT7M11 =  SUM(ISNULL(U_BiGwa02,0)),
                       BT8M11 =  SUM(ISNULL(U_BiGwa07,0))
                FROM   [@ZPY343L]
                WHERE  U_LineNum = '11'
                GROUP  BY DocEntry
                ) T11,
                (
                SELECT DocM12 =  DocEntry,
                       JIGM12 =  MAX(SUBSTRING(U_JIGDATE, 1, 6)),
                       PAYM12 =  SUM(U_GwaPay) ,
                       BNSM12 =  SUM(U_GwaBns) ,
                       INJB12 =  SUM(U_InJBns) ,
                       SUMM12 =  SUM(U_GwaSee) ,
                       BT3M12 =  SUM(ISNULL(U_BiGwa03,0) + ISNULL(U_BIGM01,0)+ ISNULL(U_BIGM02,0)+ ISNULL(U_BIGM03,0)),
                       BT1M12 =  SUM(ISNULL(U_BiGwa01,0) + ISNULL(U_BIGO01,0)),
                       BT2M12 =  SUM(U_BiGwa02 + ISNULL(U_BiGwa07,0)),
                       GABM12 =  SUM(U_GabGun),
                       JUMM12 =  SUM(U_JuMin) ,
                       KUKM12 =  SUM(U_KukAmt),
                       MEDM12 =  SUM(U_MedAmt + ISNULL(U_NGYAMT,0)),
                       GBHM12 =  SUM(U_GBHAMT),
                       JUSM12 =  SUM(U_JUSBNS),
                       GBUM12 =  SUM(ISNULL(U_GbuAmt,0)) ,
                       BU3M12 =  SUM(ISNULL(U_BiGwu03,0) + ISNULL(U_BIGX01,0)),
                       BT4M12 =  SUM(ISNULL(U_BiGwa04,0) + ISNULL(U_BIGG01,0) + ISNULL(U_BIGH01,0) + ISNULL(U_BIGH11,0) 
                                    + ISNULL(U_BIGH12,0) + ISNULL(U_BIGH13,0) + ISNULL(U_BIGI01,0) + ISNULL(U_BIGK01,0) 
                                    + ISNULL(U_BIGS01,0) + ISNULL(U_BIGT01,0) + ISNULL(U_BIGY01,0) + ISNULL(U_BIGY02,0) 
                                    + ISNULL(U_BIGY03,0) + ISNULL(U_BIGY20,0) + ISNULL(U_BIGY21,0) + ISNULL(U_BIGZ01,0)),
                       BT5M12 =  SUM(ISNULL(U_BiGwa05,0) + ISNULL(U_BIGH06,0) + ISNULL(U_BIGH07,0) + ISNULL(U_BIGH08,0) + ISNULL(U_BIGH09,0) + ISNULL(U_BIGH10,0)),
                       BT6M12 =  SUM(ISNULL(U_BiGwa06,0) + ISNULL(U_BIGQ01,0)),
                       URIB12 =  SUM(ISNULL(U_URIBNS,0)),
                       BT7M12 =  SUM(ISNULL(U_BiGwa02,0)),
                       BT8M12 =  SUM(ISNULL(U_BiGwa07,0))
                FROM   [@ZPY343L]
                WHERE  U_LineNum = '12'
                GROUP  BY DocEntry
                ) T12,
                (
                SELECT DocM13 =  DocEntry,
                       JIGM13 =  '     ',
                       PAYM13 =  SUM(U_GwaPay) ,
                       BNSM13 =  SUM(U_GwaBns) ,
                       INJB13 =  SUM(U_InJBns) ,
                       SUMM13 =  SUM(U_GwaSee) ,
                       BT3M13 =  SUM(ISNULL(U_BiGwa03,0) + ISNULL(U_BIGM01,0)+ ISNULL(U_BIGM02,0)+ ISNULL(U_BIGM03,0)),
                       BT1M13 =  SUM(ISNULL(U_BiGwa01,0) + ISNULL(U_BIGO01,0)),
                       BT2M13 =  SUM(U_BiGwa02 + ISNULL(U_BiGwa07,0)),
                       GABM13 =  SUM(U_GabGun),
                       JUMM13 =  SUM(U_JuMin) ,
                       KUKM13 =  SUM(U_KukAmt),
                       MEDM13 =  SUM(U_MedAmt + ISNULL(U_NGYAMT,0))  ,
                       GBHM13 =  SUM(U_GBHAMT),
                       JUSM13 =  SUM(U_JUSBNS),
                       GBUM13 =  SUM(ISNULL(U_GbuAmt,0)) ,
                       BU3M13 =  SUM(ISNULL(U_BiGwu03,0) + ISNULL(U_BIGX01,0)),
                       BT4M13 =  SUM(ISNULL(U_BiGwa04,0) + ISNULL(U_BIGG01,0) + ISNULL(U_BIGH01,0) + ISNULL(U_BIGH11,0) 
                                    + ISNULL(U_BIGH12,0) + ISNULL(U_BIGH13,0) + ISNULL(U_BIGI01,0) + ISNULL(U_BIGK01,0) 
                                    + ISNULL(U_BIGS01,0) + ISNULL(U_BIGT01,0) + ISNULL(U_BIGY01,0) + ISNULL(U_BIGY02,0) 
                                    + ISNULL(U_BIGY03,0) + ISNULL(U_BIGY20,0) + ISNULL(U_BIGY21,0) + ISNULL(U_BIGZ01,0)),
                       BT5M13 =  SUM(ISNULL(U_BiGwa05,0) + ISNULL(U_BIGH06,0) + ISNULL(U_BIGH07,0) + ISNULL(U_BIGH08,0) + ISNULL(U_BIGH09,0) + ISNULL(U_BIGH10,0)),
                       BT6M13 =  SUM(ISNULL(U_BiGwa06,0) + ISNULL(U_BIGQ01,0)),
                       URIB13 =  SUM(ISNULL(U_URIBNS,0)),
                       BT7M13 =  SUM(ISNULL(U_BiGwa02,0)),
                       BT8M13 =  SUM(ISNULL(U_BiGwa07,0))
                FROM   [@ZPY343L]
                WHERE  U_LineNum = '13'
                GROUP  BY DocEntry
                ) T13
        WHERE   T01.DocEntry = T02.DocM02
        AND     T01.DocEntry = T03.DocM03
        AND     T01.DocEntry = T04.DocM04
        AND     T01.DocEntry = T05.DocM05
        AND     T01.DocEntry = T06.DocM06
        AND     T01.DocEntry = T07.DocM07
        AND     T01.DocEntry = T08.DocM08
        AND     T01.DocEntry = T09.DocM09
        AND     T01.DocEntry = T10.DocM10
        AND     T01.DocEntry = T11.DocM11
        AND     T01.DocEntry = T12.DocM12
        AND     T01.DocEntry = T13.DocM13
        ) T0
        INNER JOIN  [@ZPY343H] T1 ON T0.DocEntry = T1.DocEntry
                        INNER JOIN [@PH_PY001A] T2 ON T1.U_MstCode = T2.Code
                        --INNER JOIN [@ZPY504H] T3 ON T1.U_JsnYear = T3.U_JSNYER AND T1.U_MstCode = T3.U_MSTCOD 
                        LEFT JOIN [@PH_PY005A] T4 ON T1.U_CLTCOD = T4.U_CLTCode
                       
    WHERE   T1.U_JsnYear    = @JSNYER
    AND     T1.U_CLTCOD  LIKE @CLTCOD
    AND     T1.U_DptCode LIKE @MSTDPT
    AND     T1.U_MstCode LIKE @MSTCOD
    
    /*
    AND     T2.Status LIKE CASE @JOBGBN WHEN '1' THEN '1' 
                                        WHEN '2' THEN '4'
                                        ELSE '%' END
   */                                       
    ORDER BY  T1.U_MstName,  T1.U_MstCode


-- <3. 정산구분에 따라 처리 대상 구분>ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ                        
    IF (@JOBGBN = '1')      --1연말정산(퇴사자제외)
        BEGIN
            DELETE FROM [#RPY503] 
            WHERE MSTCOD  IN (SELECT ISNULL(U_MSTCOD,'') COLLATE Korean_Wansung_CI_AS AS U_MSTCOD  
                            FROM OHEM 
                            WHERE ISNULL(CONVERT(CHAR(8), TermDate, 112),'') <> ''
                            AND   ISNULL(CONVERT(CHAR(6), TermDate, 112),'') <  @JSNYER + @STRMON
                            )
        END

    IF (@JOBGBN = '2')      --중도정산(퇴사자만)
        BEGIN
            DELETE FROM [#RPY503] 
            WHERE MSTCOD IN (SELECT ISNULL(U_MSTCOD,'') COLLATE Korean_Wansung_CI_AS AS U_MSTCOD 
                            FROM OHEM 
                            WHERE (ISNULL(CONVERT(CHAR(8), TermDate, 20),'') = ''  OR 
                                   ISNULL(CONVERT(CHAR(6), TermDate, 112),'') NOT BETWEEN @JSNYER + @STRMON AND @JSNYER + @ENDMON)
                            )
        
        END 


    SELECT * FROM [#RPY503] 
--THE END /////////////////////////////////////////////////////////////////////////////////////////////////
SET NOCOUNT OFF

--go
--Exec RPY503 '2013', '01', '12', '3', '%','%', '%'