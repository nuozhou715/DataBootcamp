select * from employees;

select last_name, count(emp_no) as "Share"
from employees
group by last_name;