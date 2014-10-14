-- temporary store procedue for remove foreign key
CREATE or REPLACE PROCEDURE tmp_dbmodelexcel_drop_table_fk(
    a_table_name IN VARCHAR2
) IS
  v_fk_name varchar2(250);
  CURSOR c_fk IS
    SELECT UC.constraint_name
      FROM user_constraints UC
     WHERE lower(UC.table_name) = lower(a_table_name)
       AND UC.constraint_type = 'R';
BEGIN

  OPEN c_fk;
  LOOP
    FETCH c_fk INTO v_fk_name;
    EXIT WHEN c_fk%NOTFOUND;
      EXECUTE IMMEDIATE 'ALTER TABLE ' || a_table_name || ' DROP CONSTRAINT ' || v_fk_name;
    END LOOP;
  CLOSE c_fk;
END tmp_dbmodelexcel_drop_table_fk;
/
-- Remove ALLDATATYPES foreign key constraint
CALL tmp_dbmodelexcel_drop_table_fk('ALLDATATYPES');

-- Remove DEPARTMENTS foreign key constraint
CALL tmp_dbmodelexcel_drop_table_fk('DEPARTMENTS');

-- Remove EMPLOYEES foreign key constraint
CALL tmp_dbmodelexcel_drop_table_fk('EMPLOYEES');

-- Remove ITEMBRANCHES foreign key constraint
CALL tmp_dbmodelexcel_drop_table_fk('ITEMBRANCHES');

-- Remove ITEMS foreign key constraint
CALL tmp_dbmodelexcel_drop_table_fk('ITEMS');

-- Remove TESTFOREIGNKEYOPTIONS foreign key constraint
CALL tmp_dbmodelexcel_drop_table_fk('TESTFOREIGNKEYOPTIONS');

-- Remove TESTFOREIGNKEYOPTIONS2 foreign key constraint
CALL tmp_dbmodelexcel_drop_table_fk('TESTFOREIGNKEYOPTIONS2');

-- Remove TESTFOREIGNKEYOPTIONS3 foreign key constraint
CALL tmp_dbmodelexcel_drop_table_fk('TESTFOREIGNKEYOPTIONS3');

-- Remove ZIPCODES foreign key constraint
CALL tmp_dbmodelexcel_drop_table_fk('ZIPCODES');

-- Remove temporary store procedue for remove foreign key
DROP PROCEDURE tmp_dbmodelexcel_drop_table_fk;

DECLARE
  v_table_is_exists integer;
BEGIN
  SELECT COUNT(*) INTO v_table_is_exists
  FROM user_tables
  WHERE lower(table_name) = lower('ALLDATATYPES');
  IF v_table_is_exists != 0 THEN
    execute immediate 'DROP TABLE ALLDATATYPES PURGE';
  END IF;

  SELECT COUNT(*) INTO v_table_is_exists
  FROM user_tables
  WHERE lower(table_name) = lower('DEPARTMENTS');
  IF v_table_is_exists != 0 THEN
    execute immediate 'DROP TABLE DEPARTMENTS PURGE';
  END IF;

  SELECT COUNT(*) INTO v_table_is_exists
  FROM user_tables
  WHERE lower(table_name) = lower('EMPLOYEES');
  IF v_table_is_exists != 0 THEN
    execute immediate 'DROP TABLE EMPLOYEES PURGE';
  END IF;

  SELECT COUNT(*) INTO v_table_is_exists
  FROM user_tables
  WHERE lower(table_name) = lower('ITEMBRANCHES');
  IF v_table_is_exists != 0 THEN
    execute immediate 'DROP TABLE ITEMBRANCHES PURGE';
  END IF;

  SELECT COUNT(*) INTO v_table_is_exists
  FROM user_tables
  WHERE lower(table_name) = lower('ITEMS');
  IF v_table_is_exists != 0 THEN
    execute immediate 'DROP TABLE ITEMS PURGE';
  END IF;

  SELECT COUNT(*) INTO v_table_is_exists
  FROM user_tables
  WHERE lower(table_name) = lower('TESTFOREIGNKEYOPTIONS');
  IF v_table_is_exists != 0 THEN
    execute immediate 'DROP TABLE TESTFOREIGNKEYOPTIONS PURGE';
  END IF;

  SELECT COUNT(*) INTO v_table_is_exists
  FROM user_tables
  WHERE lower(table_name) = lower('TESTFOREIGNKEYOPTIONS2');
  IF v_table_is_exists != 0 THEN
    execute immediate 'DROP TABLE TESTFOREIGNKEYOPTIONS2 PURGE';
  END IF;

  SELECT COUNT(*) INTO v_table_is_exists
  FROM user_tables
  WHERE lower(table_name) = lower('TESTFOREIGNKEYOPTIONS3');
  IF v_table_is_exists != 0 THEN
    execute immediate 'DROP TABLE TESTFOREIGNKEYOPTIONS3 PURGE';
  END IF;

  SELECT COUNT(*) INTO v_table_is_exists
  FROM user_tables
  WHERE lower(table_name) = lower('ZIPCODES');
  IF v_table_is_exists != 0 THEN
    execute immediate 'DROP TABLE ZIPCODES PURGE';
  END IF;

END;
/

--------------------------------
-- Create table: 'ALLDATATYPES'
--------------------------------
CREATE TABLE ALLDATATYPES (
   DATATYPEID number NOT NULL
  ,DATATYPENAME nvarchar2(15) NOT NULL
  ,DTNUMERIC number DEFAULT (1.1) NOT NULL
  ,DTNUMERIC_8_2 number(8, 2) DEFAULT (2.2) NOT NULL
  ,DTINT number DEFAULT (3) NOT NULL
  ,DTFLOAT float DEFAULT (4) NOT NULL
  ,DTFLOAT_10 float(10) DEFAULT (5) NOT NULL
  ,DTBINARYFLOAT binary_float DEFAULT (6) NOT NULL
  ,DTBINARYDOUBLE binary_double DEFAULT (7) NOT NULL
  ,DTDATE date NOT NULL
  ,DTTIMESTAMP timestamp(6) NOT NULL
  ,DTTIMESTAMPWITHZONE timestamp(6) with time zone NOT NULL
  ,DTTIMESTAMPWITHLOCAL timestamp(6) with local time zone NOT NULL
  ,DTINTERVALYEARTOMONTH interval year(2) to month NOT NULL
  ,DTINTERVALDAYTOSECOND interval day(2) to second(6) NOT NULL
  ,DTCHAR char(255) NULL
  ,DTVARCHAR varchar2(255) DEFAULT ('') NOT NULL
  ,DTCLOB clob NULL
  ,DTNCHAR nchar(255) NULL
  ,DTNVARCHAR2 nvarchar2(255) DEFAULT ('A') NOT NULL
  ,DTNCLOB nclob NULL
  ,DTRAW raw(1000) NULL
  ,DTLONGRAW long raw NULL
  ,DTBLOB blob NULL
  ,DTBFINE bfile NULL
);
ALTER TABLE ALLDATATYPES ADD CONSTRAINT PK_ALLDATATYPES PRIMARY KEY (DATATYPEID);
ALTER TABLE ALLDATATYPES ADD CONSTRAINT IK_ALLDATATYPES_DATATYPENAME UNIQUE (DATATYPENAME);
COMMENT ON TABLE ALLDATATYPES IS 'Sample table for most common data types'' definitions.';
COMMENT ON COLUMN ALLDATATYPES.DATATYPEID IS 'Test single'' quatation (Label: Data Type ID)';
COMMENT ON COLUMN ALLDATATYPES.DATATYPENAME IS 'Test single quatation (Label: Data Type''s Name)';
COMMENT ON COLUMN ALLDATATYPES.DTNUMERIC IS 'Numeric Data Types';
COMMENT ON COLUMN ALLDATATYPES.DTNUMERIC_8_2 IS 'Numeric Data Types';
COMMENT ON COLUMN ALLDATATYPES.DTINT IS 'Numeric Data Types';
COMMENT ON COLUMN ALLDATATYPES.DTFLOAT IS 'Numeric Data Types';
COMMENT ON COLUMN ALLDATATYPES.DTFLOAT_10 IS 'Numeric Data Types';
COMMENT ON COLUMN ALLDATATYPES.DTBINARYFLOAT IS 'Floating point number';
COMMENT ON COLUMN ALLDATATYPES.DTBINARYDOUBLE IS 'Floating point number';
COMMENT ON COLUMN ALLDATATYPES.DTDATE IS 'Date and Time';
COMMENT ON COLUMN ALLDATATYPES.DTTIMESTAMP IS 'Date and Time';
COMMENT ON COLUMN ALLDATATYPES.DTTIMESTAMPWITHZONE IS 'Date and Time';
COMMENT ON COLUMN ALLDATATYPES.DTTIMESTAMPWITHLOCAL IS 'Date and Time';
COMMENT ON COLUMN ALLDATATYPES.DTINTERVALYEARTOMONTH IS 'Date and Time';
COMMENT ON COLUMN ALLDATATYPES.DTINTERVALDAYTOSECOND IS 'Date and Time';
COMMENT ON COLUMN ALLDATATYPES.DTCHAR IS 'Character Data Types';
COMMENT ON COLUMN ALLDATATYPES.DTVARCHAR IS 'Character Data Types';
COMMENT ON COLUMN ALLDATATYPES.DTCLOB IS 'Character Data Types';
COMMENT ON COLUMN ALLDATATYPES.DTNCHAR IS 'Character Data Types';
COMMENT ON COLUMN ALLDATATYPES.DTNVARCHAR2 IS 'Character Data Types';
COMMENT ON COLUMN ALLDATATYPES.DTNCLOB IS 'Character Data Types';
COMMENT ON COLUMN ALLDATATYPES.DTRAW IS 'Large Object Data Types';
COMMENT ON COLUMN ALLDATATYPES.DTLONGRAW IS 'Large Object Data Types';
COMMENT ON COLUMN ALLDATATYPES.DTBLOB IS 'Large Object Data Types';
COMMENT ON COLUMN ALLDATATYPES.DTBFINE IS 'Large Object Data Types';

--------------------------------
-- Create table: 'DEPARTMENTS'
--------------------------------
CREATE TABLE DEPARTMENTS (
   DEPARTMENTID number NOT NULL
  ,DEPARTMENTNAME nvarchar2(50) NOT NULL
  ,PARENTID number NULL
  ,MANAGERID number NULL
);
ALTER TABLE DEPARTMENTS ADD CONSTRAINT PK_DEPARTMENTS PRIMARY KEY (DEPARTMENTID);
COMMENT ON TABLE DEPARTMENTS IS 'The department table.';
COMMENT ON COLUMN DEPARTMENTS.DEPARTMENTID IS '(Label: Department ID)';
COMMENT ON COLUMN DEPARTMENTS.DEPARTMENTNAME IS '(Label: Department Name)';
COMMENT ON COLUMN DEPARTMENTS.PARENTID IS '(Label: Parent Department)';
COMMENT ON COLUMN DEPARTMENTS.MANAGERID IS '(Label: Manager)';

--------------------------------
-- Create table: 'EMPLOYEES'
--------------------------------
CREATE TABLE EMPLOYEES (
   EMPLOYEEID number NOT NULL
  ,LASTNAME nvarchar2(50) NOT NULL
  ,FIRSTNAME nvarchar2(50) NOT NULL
  ,DEPARTMENTID number NOT NULL
);
ALTER TABLE EMPLOYEES ADD CONSTRAINT PK_EMPLOYEES PRIMARY KEY (EMPLOYEEID);
CREATE INDEX IK_EMPLOYEES_FIRSTNAME_LAS0000 ON EMPLOYEES (FIRSTNAME, LASTNAME);
CREATE INDEX IK_EMPLOYEES_LASTNAME ON EMPLOYEES (LASTNAME);
COMMENT ON TABLE EMPLOYEES IS 'Employees';
COMMENT ON COLUMN EMPLOYEES.EMPLOYEEID IS '(Label: EmployeeID)';
COMMENT ON COLUMN EMPLOYEES.LASTNAME IS '(Label: Last Name)';
COMMENT ON COLUMN EMPLOYEES.FIRSTNAME IS '(Label: First Name)';
COMMENT ON COLUMN EMPLOYEES.DEPARTMENTID IS '(Label: Department)';

--------------------------------
-- Create table: 'ITEMBRANCHES'
--------------------------------
CREATE TABLE ITEMBRANCHES (
   ITEMID number NOT NULL
  ,SUBITEMID number NOT NULL
  ,BRANCHID number NOT NULL
  ,ITEMVALUE nvarchar2(255) NOT NULL
);
ALTER TABLE ITEMBRANCHES ADD CONSTRAINT PK_ITEMBRANCHES PRIMARY KEY (ITEMID, SUBITEMID, BRANCHID);
CREATE INDEX IK_ITEMBRANCHES_ITEMID_SUB0000 ON ITEMBRANCHES (ITEMID, SUBITEMID);
CREATE INDEX IK_ITEMBRANCHES_SUBITEMID_0000 ON ITEMBRANCHES (SUBITEMID, ITEMID);
COMMENT ON TABLE ITEMBRANCHES IS '';

--------------------------------
-- Create table: 'ITEMS'
--------------------------------
CREATE TABLE ITEMS (
   ITEMID number NOT NULL
  ,SUBITEMID number NOT NULL
  ,ITEMNAME nvarchar2(255) NOT NULL
);
ALTER TABLE ITEMS ADD CONSTRAINT PK_ITEMS PRIMARY KEY (ITEMID, SUBITEMID);
COMMENT ON TABLE ITEMS IS '';

--------------------------------
-- Create table: 'TESTFOREIGNKEYOPTIONS'
--------------------------------
CREATE TABLE TESTFOREIGNKEYOPTIONS (
   DEPARTMENTID number NOT NULL
  ,MEMO nvarchar2(50) NOT NULL
);
ALTER TABLE TESTFOREIGNKEYOPTIONS ADD CONSTRAINT PK_TESTFOREIGNKEYOPTIONS PRIMARY KEY (DEPARTMENTID);
COMMENT ON TABLE TESTFOREIGNKEYOPTIONS IS 'Test ForeignKey Actions: CASCADE/NULL/No Action';

--------------------------------
-- Create table: 'TESTFOREIGNKEYOPTIONS2'
--------------------------------
CREATE TABLE TESTFOREIGNKEYOPTIONS2 (
   OPTIONID number NOT NULL
  ,DEPARTMENTID number NULL
  ,MEMO nvarchar2(50) NOT NULL
);
ALTER TABLE TESTFOREIGNKEYOPTIONS2 ADD CONSTRAINT PK_TESTFOREIGNKEYOPTIONS2 PRIMARY KEY (OPTIONID);
COMMENT ON TABLE TESTFOREIGNKEYOPTIONS2 IS 'Test ForeignKey Actions: CASCADE/NULL/No Action';

--------------------------------
-- Create table: 'TESTFOREIGNKEYOPTIONS3'
--------------------------------
CREATE TABLE TESTFOREIGNKEYOPTIONS3 (
   OPTIONID number NOT NULL
  ,DEPARTMENTID number NULL
  ,MEMO nvarchar2(50) NOT NULL
);
ALTER TABLE TESTFOREIGNKEYOPTIONS3 ADD CONSTRAINT PK_TESTFOREIGNKEYOPTIONS3 PRIMARY KEY (OPTIONID);
COMMENT ON TABLE TESTFOREIGNKEYOPTIONS3 IS 'Test ForeignKey Actions: CASCADE/NULL/No Action';

--------------------------------
-- Create table: 'ZIPCODES'
--------------------------------
CREATE TABLE ZIPCODES (
   ZIPCODE varchar2(8) NOT NULL
  ,ADDRESS1 nvarchar2(255) NOT NULL
  ,ADDRESS2 nvarchar2(255) DEFAULT ('') NOT NULL
  ,ADDRESS3 nvarchar2(255) DEFAULT ('') NOT NULL
);
CREATE INDEX IK_ZIPCODES_ZIPCODE ON ZIPCODES (ZIPCODE);
COMMENT ON TABLE ZIPCODES IS 'Zip codes';
COMMENT ON COLUMN ZIPCODES.ZIPCODE IS 'Zip code is not unique. (Label: Zip Code)';
COMMENT ON COLUMN ZIPCODES.ADDRESS1 IS '(Label: Address1)';
COMMENT ON COLUMN ZIPCODES.ADDRESS2 IS '(Label: Address2)';
COMMENT ON COLUMN ZIPCODES.ADDRESS3 IS '(Label: Address3)';


-- Create foreign keys for table: 'DEPARTMENTS'
ALTER TABLE DEPARTMENTS
  ADD CONSTRAINT FK_DEPARTMENTS_MANAGERID
  FOREIGN KEY (MANAGERID)
  REFERENCES EMPLOYEES(EMPLOYEEID);
ALTER TABLE DEPARTMENTS
  ADD CONSTRAINT FK_DEPARTMENTS_PARENTID
  FOREIGN KEY (PARENTID)
  REFERENCES DEPARTMENTS(DEPARTMENTID);

-- Create foreign keys for table: 'EMPLOYEES'
ALTER TABLE EMPLOYEES
  ADD CONSTRAINT FK_EMPLOYEES_DEPARTMENTID
  FOREIGN KEY (DEPARTMENTID)
  REFERENCES DEPARTMENTS(DEPARTMENTID);

-- Create foreign keys for table: 'ITEMBRANCHES'
ALTER TABLE ITEMBRANCHES
  ADD CONSTRAINT FK_ITEMBRANCHES_ITEMID_SUB0000
  FOREIGN KEY (ITEMID,SUBITEMID)
  REFERENCES ITEMS(ITEMID,SUBITEMID) ON DELETE CASCADE;

-- Create foreign keys for table: 'TESTFOREIGNKEYOPTIONS'
ALTER TABLE TESTFOREIGNKEYOPTIONS
  ADD CONSTRAINT FK_TESTFOREIGNKEYOPTIONS_D0000
  FOREIGN KEY (DEPARTMENTID)
  REFERENCES DEPARTMENTS(DEPARTMENTID);

-- Create foreign keys for table: 'TESTFOREIGNKEYOPTIONS2'
ALTER TABLE TESTFOREIGNKEYOPTIONS2
  ADD CONSTRAINT FK_TESTFOREIGNKEYOPTIONS2_0000
  FOREIGN KEY (DEPARTMENTID)
  REFERENCES DEPARTMENTS(DEPARTMENTID) ON DELETE CASCADE;

-- Create foreign keys for table: 'TESTFOREIGNKEYOPTIONS3'
ALTER TABLE TESTFOREIGNKEYOPTIONS3
  ADD CONSTRAINT FK_TESTFOREIGNKEYOPTIONS3_0000
  FOREIGN KEY (DEPARTMENTID)
  REFERENCES DEPARTMENTS(DEPARTMENTID) ON DELETE SET NULL;

