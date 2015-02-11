package watsondb;

import java.util.Collection;
import java.util.Hashtable;

public class Dashboard {
	
	private Hashtable<String, Department> departments;
	private Hashtable<String, Person> faculty;
	
	public Dashboard() {
		departments = new Hashtable<String, Department>();
		faculty = new Hashtable<String, Person>();
	}
	
	public void addNewDepartment(String deptName) {
		departments.put(deptName, new Department(deptName));
	}
	
	public Department getDepartment(String deptName) {
		if (departments.containsKey(deptName)) {
			return departments.get(deptName);
		} else {
			return null;
		}
	}
	
	public Collection<Department> getDepartments() {
		return departments.values();
	}
	
	public boolean hasDepartment(String deptName) {
		if (departments.containsKey(deptName))
			return true;
		else
			return false;
	}
	
	public void printDepartments() {
		Collection<Department> values = departments.values();
		for (Department department : values) {
			System.out.println(department);
		}
	}
	
	public void addNewPerson(String personName, Department department) {
		faculty.put(personName, new Person(personName, department));
	}
	
	public Person getPerson(String personName) {
		if (faculty.containsKey(personName)) {
			return faculty.get(personName);
		} else {
			return null;
		}
	}
	
	public Collection<Person> getFaculty() {
		return faculty.values();
	}
	
	public boolean hasPerson(String personName) {
		if (faculty.containsKey(personName))
			return true;
		else
			return false;
	}
	
	public void printFaculty() {
		Collection<Person> values = faculty.values();
		for (Person person : values) {
			System.out.println(person);
		}
	}
}
