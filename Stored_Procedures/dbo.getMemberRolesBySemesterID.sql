CREATE procedure [dbo].sp_getMemberRoleBySemesterID
@semesterID nvarchar (75)

AS

Select * from MEMBER_ROLE
where SemesterID=@semesterID
RETURN 0