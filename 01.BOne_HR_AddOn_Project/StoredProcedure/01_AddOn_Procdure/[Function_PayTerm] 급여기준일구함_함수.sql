CREATE	FUNCTION Func_PAYTerm  (@JobDate NVarchar(10), @PAYSEL Nvarchar(8))

RETURNS nvarchar(43)
--WITH Encryption
AS
	/*==========================================================================================
		�Լ���		    : Func_PAYTerm
		�Լ�����	    : �޿��ͼӿ��� ������, �����ϱ���
		������			: �Թ̰�
		�۾�����		: 2007-01-11
		�۾�������		: �Թ̰�
		�۾���������	: 2007-01-11
		�۾�����		: �޿��۾����� �ͼ����ڸ� ��ȸ��
		�۾�����		: 
	===========================================================================================*/
	-- DROP FUNCTION Func_PAYTerm
	-- SELECT DBO.Func_PAYTerm('2008-11-21', '1')
	-- SELECT DBO.Func_PAYTerm('2007-02-01', '1') 

BEGIN
	DECLARE	@U_STRMON		 AS SmallInt
	DECLARE	@U_STRDAY		 AS Nvarchar(2)
	DECLARE	@U_JIGMON		 AS SmallInt
	DECLARE	@U_JIGDAY		 AS Nvarchar(2)
	DECLARE	@U_BNSEMM		 AS SmallInt
	DECLARE	@U_BNSEDD		 AS Nvarchar(2)
	
	DECLARE @FrDate NvarChar(8)
    	   ,@ToDate  NvarChar(8)
    	   ,@JIGBIL  NvarChar(8)
    	   ,@StDate  NvarChar(8)
    	   ,@BNSLMT  NvarChar(8)

--1.�޿������� ���� 
SELECT TOP 1	@U_STRMON = T0.U_STRMON, 
		@U_STRDAY = T0.U_STRDAY, 
		@U_JIGMON = T0.U_JIGMON, 
		@U_JIGDAY = T0.U_JIGDAY, 
		@U_BNSEMM = T0.U_BNSEMM, 
		@U_BNSEDD = T0.U_BNSEDD 
FROM [@PH_PY107B] T0 
WHERE 	T0.U_PAYSEL = @PAYSEL
AND	T0.Code <= SUBSTRING(@JobDate, 1, 4) +  SUBSTRING(@JobDate, 6, 2)
ORDER BY T0.Code Desc

SET @U_STRDAY = CASE WHEN @U_STRDAY <10 THEN '0' + @U_STRDAY ELSE @U_STRDAY END
SET @U_JIGDAY = CASE WHEN @U_JIGDAY <10 THEN '0' + @U_JIGDAY ELSE @U_JIGDAY END
SET @U_BNSEDD = CASE WHEN @U_BNSEDD <10 THEN '0' + @U_BNSEDD ELSE @U_BNSEDD END


IF ISNULL(@U_STRDAY, '') <> '' 
	BEGIN
	--2.�޿� �ͼ� ������
	    SET @FrDate = CONVERT(CHAR(8), DATEADD(Month, @U_STRMON,@JobDate), 112)
	    SET @FrDate = SUBSTRING(@FrDate, 1, 6) + @U_STRDAY
	--3.�޿� �ͼ� ������
	   SET @ToDate = CONVERT(CHAR(8),DATEADD(Day, -1, DATEADD(Month, 1,@FrDate)), 112)
	--4.�޿� ������
	    IF @U_JIGDAY = 0  --����
	        BEGIN
	        SET @JIGBIL = CONVERT(CHAR(8), DATEADD(Month,@U_JIGMON, @JobDate), 112)
	        
	        SET @StDate = CONVERT(CHAR(8), DATEADD(Day, -1 , DATEADD(Month,1,@JIGBIL)), 112)  --������-�Ϸ�
	        SET @U_JIGDAY = CAST(Day(@StDate) as nvarchar)
	        END
	    ELSE
	        BEGIN
	       	SET @JIGBIL =CONVERT(CHAR(8), DATEADD(Month,@U_JIGMON, @JobDate), 112)
	        
		END 
		SET @JIGBIL =SUBSTRING(@JIGBIL , 1, 6) + @U_JIGDAY
	--5.��������
	    IF @U_BNSEDD = 0  --����
	        BEGIN
	        SET @BNSLMT = CONVERT(CHAR(8), DATEADD(Month, @U_BNSEMM, @JobDate), 112)
	        
	        SET @StDate = CONVERT(CHAR(8), DATEADD(Day, -1 , DATEADD(Month,1,@BNSLMT)), 112)  --������-�Ϸ�
	        SET @U_BNSEDD = CAST(Day(@StDate) as nvarchar)
	        END
	    ELSE
	        BEGIN
	       	SET @BNSLMT =CONVERT(CHAR(8), DATEADD(Month,@U_BNSEMM, @JobDate), 112)
	        
		END 
		SET @BNSLMT =SUBSTRING(@BNSLMT , 1, 6) + @U_BNSEDD
		

	END 

--    RETURN(CONVERT(CHAR(8),@FrDate,112)+CONVERT(CHAR(8),@ToDate,112)+CONVERT(CHAR(8),@JIGBIL,112))

  RETURN(@FrDate + ' ' + @ToDate + ' ' + @JIGBIL + ' ' + @BNSLMT)
END


