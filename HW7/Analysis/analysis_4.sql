select * from dept_emp;
select * from departments;
select * from employees;

select dept_emp.dept_no, employees.last_name, employees.first_name, departments.dept_name
from dept_emp
left join employees
on employees.emp_no = dept_emp.emp_no
left join departments
on departments.dept_no = dept_emp.dept_no;