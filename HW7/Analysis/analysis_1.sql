select employees.emp_no, employees.gender, salaries.emp_no, salaries.salary
from employees
full join salaries
on employees.emp_no = salaries.emp_no;

select count(*) from (select employees.emp_no, employees.gender, salaries.salary
from employees
full join salaries
on employees.emp_no = salaries.emp_no) as train where train.emp_no is NULL;