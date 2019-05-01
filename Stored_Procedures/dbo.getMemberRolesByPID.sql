CREATE procedure [dbo].sp_getMemberRoleByPID
@pID nvarchar (7)

AS

Select * from MEMBER_ROLE
where PID=@pID
RETURN 0