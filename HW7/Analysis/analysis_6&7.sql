select * from dept_emp;
select * from employees;
select * from departments;

-- Analysis 6
select dept_emp.emp_no, employees.last_name, employees.first_name, departments.dept_name
from (dept_emp
	left join employees
	on employees.emp_no = dept_emp.emp_no
	left join departments
	on departments.dept_no = dept_emp.dept_no
)
where departments.dept_name = 'Sales';

-- Analysis 7
select dept_emp.emp_no, employees.last_name, employees.first_name, departments.dept_name
from (dept_emp
	left join employees
	on employees.emp_no = dept_emp.emp_no
	left join departments
	on departments.dept_no = dept_emp.dept_no
)
where (
	departments.dept_name = 'Sales' or departments.dept_name = 'Development'
	);