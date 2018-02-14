set ANSI_NULLS ON
set QUOTED_IDENTIFIER ON
go



/****************************************************************************************************************/
/*  Module         : 생산관리																				    */
/*  Description    : 휘팅이동등록																				*/
/*  ALTER  Date    : 2010.10.22																					*/
/*  Modified Date  :																							*/
/*  Creator        : Youn Je Hyung                                                                              */
/*  Company        : Poongsan Holdings																			*/
/****************************************************************************************************************/
--CREATE  PROC [dbo].[PS_PP075_01]
ALTER     PROC [dbo].[PS_PP075_01]
(
  @BaseDate				as datetime
 )
AS
SET NOCOUNT ON
--BEGIN /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
declare @_MovDocNo	nvarchar(11),
		@_Cnt		int,
		@_SeqNo		int,
		@_MaxNo		nvarchar(3)

select @_Cnt=count(U_MovDocNo) from [@PS_PP075H]
where left(U_MovDocNo,8) = @BaseDate

--select @_Cnt

if @_Cnt=0 begin --해당년월일에 등록번호가 없다면
	select convert(char(8),@BaseDate,112)+'001'

end else begin --해당년월일에 등록번호가 있다면
	select @_MovDocNo=max(U_MovDocNo) from [@PS_PP075H]
	where left(U_MovDocNo,8) = @BaseDate
	
	set @_SeqNo= convert(int,right(@_MovDocNo,3)+1)
	
	if len(@_SeqNo)=1 begin 
		set @_MaxNo='00'+convert(char(1),@_SeqNo) 
	end else if len(@_SeqNo)=2 begin 
		set @_MaxNo='0'+convert(char(2),@_SeqNo) 
	end else begin
		set @_MaxNo=convert(char(3),@_SeqNo) 
	end
	
	select convert(char(8),@BaseDate,112)+@_MaxNo
end


-----------------------------------------------------------------------------------------------------------------------------------------
--EXEC PS_PP075_01 '20101022'
--EXEC PS_PP075_01 '20101023'




	