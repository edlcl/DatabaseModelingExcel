SET SERVEROUTPUT ON SIZE 100000
WHENEVER SQLERROR EXIT SQL.SQLCODE ROLLBACK
WHENEVER OSERROR EXIT -1

define user_name=testuser
define user_password=1
define dbf_location=C:\ORA_TEST
define tablespace_name=testtn
define log_filename=Ora_CreateDB.log

define directory_name=OP_DATAPUMP_DIRECTORY

SPOOL &log_filename
-- create tablespaces --
create tablespace &tablespace_name datafile '&dbf_location/&tablespace_name.dbf' size 512 M reuse autoextend on next 128 M maxsize 1024 M;

-- create cognos user --
create user &user_name identified by &user_password default tablespace &tablespace_name temporary tablespace TEMP;

-- grant rights to the user --
grant connect, resource to &user_name;
grant create table to &user_name;
grant create trigger to &user_name;
grant create view to &user_name;
grant create procedure to &user_name;
grant create sequence to &user_name;
grant select_catalog_role to &user_name;

-- grant unlimited access to the table spaces to the user --
alter user &user_name quota unlimited on &tablespace_name;