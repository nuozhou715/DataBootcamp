select * from employees;

select employees.emp_no, employees.gender, employees.first_name, employees.last_name, employees.hire_date, salaries.salary
from employees
full join salaries
on employees.emp_no = salaries.emp_no;

select hire.emp_no, hire.first_name, hire.last_name, hire.hire_date
from (
	select employees.emp_no, employees.gender, employees.first_name, employees.last_name, employees.hire_date, salaries.salary
	from employees
	full join salaries
	on employees.emp_no = salaries.emp_no
) as hire
where (select extract(year from hire_date)) = '1986';