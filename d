SELECT class, count(class) as cnt
FROM student group by class

##  Count * will be taking on class ignoring null 
SELECT class, count(class) as cnt, count(*) as CntStr
FROM student group by class

## Group by class and sex
SELECT class, count(class) as cnt, count(*), sex as CntStr
FROM student group by class, sex

## Class with only female
SELECT class, count(class) as cnt, count(*), sex as CntStr
FROM student where sex="F" group by class, sex



## Only female with count >= 2
SELECT class, count(class) as cnt, count(*) as CntStr
FROM student where sex="F" group by class  having count(*) >= 2


## Ordering Data
SELECT id, name, sex, mtest
FROM student order by mtest

## Order in desecding order
SELECT id, name, sex, mtest
FROM student order by mtest desc

## Order by two
SELECT id, name, sex, mtest
FROM student order by sex, mtest

## Ordering based on class 
SELECT * FROM student order by class

SELECT class, name FROM student order by class, name


SELECT class, count(class) as cnt, count(*), sex as CntStr
FROM student where sex="F" group by class, sex order by 1

## Top in excel to return the top no of rows  rownum/limit is called in oracle

SELECT top 1 class, count(class) as cnt, count(*), sex as CntStr
FROM student where sex="F" group by class, sex order by 1

## top 10
SELECT top 10 class, name
FROM student 

## Top 10 percent
SELECT top 10 percent class, name
FROM student 

## top used with Order by 
SELECT top 2 percent class, name
FROM student order by mtest desc




SELECT dcode, count(dcode) FROM student where sex = "F" group by dcode

SELECT name
FROM student 
where sex = "M" and Class ="1A" 
order by name


SELECT avg(mtest) as average, class
FROM  
 student
where sex = "M" 
group by class

SELECT name, dcode
FROM student 
where  Class ="1B" 
order by dcode

SELECT count(dcode), dcode
FROM  
 student group by dcode
order by dcode desc

SELECT name, class
FROM  
 student
where sex ="M"
order by class

SELECT top 10 percent name, mtest
FROM  
 student
where sex ="F"



SELECT TOP 1 name FROM  (
  SELECT top 2 name, mtest  FROM student
  ORDER BY mtest DESC
) AS em ORDER BY mtest ASC

###****
SELECT name, count(name) as Repeat
FROM  
 student
group by name
having count(name) = 1
###***
select count(*) from (
SELECT  name, count(name) as NameRepeat
FROM  
 student
group by name
having count(name) = 1
)

## Date
SELECT round(avg((date() - DOB)/365),0) as AvgAge
FROM student

select (month(DOB)) as month, count(*) as frequency
from student
group by month(DOB)


##### TEst functions


select name, mid(name,2,3) as midname, len(name) as len, right(name,2) as RIGHT1, left(name,2) as left1, replace(name,"a","b",1,1) as replace1, ucase(name) as UpperCase
from student


## *** Union between two tables 
select name, class from student
Union
select fullname, class from phy


## Top Queries
Q1


SELECT top 10 percent class, name
FROM student 


select * from
student
 where 
mtest=(select max(mtest) from student
    where mtest<(select max(mtest) from student));


select mtest, fullname from student where 
mtest = (
SELECT max(mtest) -1 
FROM student 
)


select top 1 fullname,mtest
from(select top 3 fullname,mtest from student order by mtest desc)
order by mtest



################# Text Functions


Q ANs
select left(name,1) from student where dcode = "YMT"

select left(name, len(name)-1) 
 from student

select mid(name,1,len(name)-1) 
 from student

select right(name,2) from student

SELECT Mid(name,(Len(name)-2),2) AS Last2CharMid
FROM student;

SELECT UCase(name) AS Expr1
FROM student;

SELECT name
FROM student
where name not like "*e*e*" and fullname like "*e*"

select fullname from student
where len(fullname) - len(replace(fullname,"e","")) = 1;

################## Union/ Intersection/Difference

select id, fullname from bridge 
Union 
select id, fullname from chess


select id, fullname from bridge 
where fullname in (
select fullname from chess )


select id, fullname from chess 


select id, fullname from bridge 
where fullname not in (
select fullname from chess )

select id, fullname from bridge
where fullname not in (
select fullname from chess )
Union
select id, fullname from chess
where fullname not in (
select fullname from bridge )


**** Exists / Not Exists
Q1
select * from student 
where
dcode = "HHM" and mtest > 80

Q1
SELECT *
  FROM student
 WHERE EXISTS 
      (SELECT *
         FROM student
        WHERE dcode = "HHM" and mtest > 80)

SELECT fullname, dcode, Mtest
FROM student
WHERE Dcode = "HHM" AND NOT EXISTS
(SELECT dcode
FROM student
WHERE dcode = "HHM" and mtest < 80) 

Q2
select fullname from 
(
SELECT  fullname, count(fullname) as NameRepeat
FROM  
 student
group by fullname
having count(fullname) = 1
)


********************** JOIN QUERIES ****************
Q25
select fullname, Type from student inner JOIN music on student.id=music.id 
order by type

Q26
select  Type, class, count(class) as CountClass
from student inner JOIN music on student.id=music.id 
where type = "Piano"
group by class, type order by type


select a.class, b.type , count(a.class)
from student a, music b
where a.id = b.id and type = "Piano"
group by a.class, type 


/********** Homework/

select a.class,  a.fullname, b.fullname
from 
student a
inner join student b on (a.class= b.class) where a.fullname <> b.fullname  and a.id < b.id
group by a.class, a.fullname, b.fullname   


select  distinct a.class , a.fullname, b.fullname
from 
student a, Student b
where a.class = b.class and a.fullname <> b.fullname
and a.id < b.id


Q27
select fullname, student.id from student
where id not in (select id from music)

Q27
SELECT student.id, student.fullname
FROM student
    LEFT JOIN music ON student.id = music.id
WHERE student.id not in (music.id)


Q28
select  student.id, student.fullname, music.Type 
from 
student 
left JOIN
music on student.id=music.id 
order by student.id
union
select student.id, student.fullname, music.Type 
from 
student 
right JOIN
music on student.id=music.id 
order by student.id


****************************************************************** 07/11/2016
**#####################################################################

select fullname, sex, class from student
where class="1A" and exists(select id from 
student where class="1A" and sex="F")

select t1.manf,t1.brand from beer t1 where 
not exists 
(select t2.manf from beer t2 where
t2.brand <>t1.brand and t1.manf=t2.manf)




********************************************************** SQL Case Study 4

Q2. 
SELECT count(booking_id) as Booking_Count,  Booking_Mode, Booking_Type
FROM Data
where booking_type = "p2p" 
group by booking_mode, booking_type

Q4

SELECT top 5 avg(fare) as Average_Revenue, drop_area
FROM Data
where drop_area <> NULL
group by drop_area
order by avg(fare) desc


Q5.
select top 5 pickup_area ,  driver_number from data 





Q7
SELECT *
FROM Data 
where confirmed_at  between #11/01/2013 00:00# and #11/07/2013 23:59#
order by confirmed_at



***************  JOINS
		
Q9 .... Inner Join
SELECT d.pickup_area, d.Booking_id, l.area,l.city_id
FROM Data d INNER JOIN Localities l on d.pickup_area = l.area


Q9. .... Left Join
SELECT d.pickup_area, d.Booking_id, l.area,l.city_id
FROM Data d Left JOIN Localities l on d.pickup_area = l.area

Q9 ... Right Join
SELECT d.pickup_area, d.Booking_id, l.area,l.city_id
FROM Data d Right JOIN Localities l on d.pickup_area = l.area


Q9 ... Union
SELECT d.pickup_area, d.Booking_id, l.area,l.city_id
FROM Data d Left JOIN Localities l on d.pickup_area = l.area
Union 
SELECT d.pickup_area, d.Booking_id, l.area,l.city_id
FROM Data d Right JOIN Localities l on d.pickup_area = l.area
		










**********************************   18/11/2016


'Data Defiantion language
'create
'we create table row, relationship
'while creating create the parent first while dropping drop the child first
create table publisher  (publisher is the table name)
(
pubid integer
pubname Char(255)/VarChar(255)/TEXT(255)

'Vchar is dynamic upto 255 depending upon the size of the data that much memony will be assign where as Char one is static 
if its not memory efficient than why we have Char and Text option because Vchar is always recaluclated it not data effecient where as vchar will be calculated each time you enter in the system. Vchar do more work.

pubCountry Varchar(100)

Constraint pub_pk Primpary key(PubID)   'pub_pk is a contraint name
Constraint pub_unq UNIQUE(PubName)

)

create table publisher 
(
pubid autoincremanet integer,
pubname Char(100) not null,
pubCountry Varchar(100),
Constraint pub_pk Primary key(PubID),  
Constraint pub_unq UNIQUE(PubName)
)


insert into Books publisher (pubid,pubname,pubCountry) values (901,"New Depo","India")

insert into publisher(pubid,pubname,pubCountry) values (902,"Top Pub","New Zealand")



create table Books
(
ISBN Char(12),
Title VarChar(100),
EdNo Integer,
Price Integer,
PID integer,
Constraint books_pk PRIMARY KEY(ISBN),
Constraint books_fk FOREIGN KEY(PID) REFERENCES publisher(PubID) 

These two will not work in ACCESS SQL but for reference
'Constraint books_CHK CHECK(Price>0)
'Constraint books_def DEFAULT Price =100

)


insert into Books
(ISBN,Title,EdNo,Price,PID) values ("AB12","STAT",35,400,901)
insert into Books
(ISBN,Title,EdNo,Price,PID) values ("AB13","GEO",36,500,902)

update books set Price = 900 where pid= 902




'alter
'chaging the column

'alter table publisher ADD/DROp/ALTER
alter table publisher ADD pubemail VarChar(100)

alter table publisher drop contraint pub_pk

alter table publisher Add contraint pub_pk primary key(PubID)

alter table tablename rename column oldname new name

'drop
'to drop the table completed
'To drop a table you need to drop a child first
drop table books  (Where books is the table name)


Data Manupulation Language

insert
insert into tablename(dilefname1,2) values ("111",111)
to allow multiple rows either import the row from table meaning exptract the data and insert the row

select id,fullname, maths from student where mtest between 90 and 95

insert into table_name_where_rows_need_to_add (id,fullname)
select id,fullname, maths from student where mtest between 90 and 95
' if you dont specify field all the values from all the variables will be inserted



update
update tablename  set fullname="Prats" where fullName = "Prateek"




delete
to delete any row

delete * from table_name where fullname like "Wendy"  '(passing the condition to remove the row from the table)



' Once done can not be undone


'Create a table with the select statement
'select id, name, mtest into Create_result_output_table from student where mtest > 80

'Update the Scholar table by reducing everyone’s marks by 10% 
update scholar set mtest = mtest - (10*mtest/100)

************************* CLASS WORK   ***********************

create table publisher 
(
pubid autoincremanet integer,
pubname Char(100) not null,
pubCountry Varchar(100),
Constraint pub_pk Primary key(PubID),  
Constraint pub_unq UNIQUE(PubName)
)


insert into Books publisher (pubid,pubname,pubCountry) values (901,"New Depo","India")

insert into publisher(pubid,pubname,pubCountry) values (902,"Top Pub","New Zealand")



create table Books
(
ISBN Char(12),
Title VarChar(100),
EdNo Integer,
Price Integer,
PID integer,
Constraint books_pk PRIMARY KEY(ISBN),
Constraint books_fk FOREIGN KEY(PID) REFERENCES publisher(PubID) 
)


insert into Books
(ISBN,Title,EdNo,Price,PID) values ("AB12","STAT",35,400,901)
insert into Books
(ISBN,Title,EdNo,Price,PID) values ("AB13","GEO",36,500,902)

update books set Price = 900 where pid= 902

delete * from books
