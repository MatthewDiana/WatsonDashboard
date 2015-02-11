package watsondb;

import java.math.BigDecimal;

public class Department {

	private String name;
	private BigDecimal proposalsTotal;
	private BigDecimal expendituresTotal;
	private BigDecimal committedFundsTotal;
	private BigDecimal stipendTotal;
	private BigDecimal awardsTotal;
	
	public Department(String name) {
		this.name = name;
		proposalsTotal = new BigDecimal("0");
		expendituresTotal = new BigDecimal("0");
		committedFundsTotal = new BigDecimal("0");
		stipendTotal = new BigDecimal("0");
		awardsTotal = new BigDecimal("0");
	}

	public String getName() {
		return name;
	}
	
	public void increaseProposalsTotal(BigDecimal amount) {
		proposalsTotal = proposalsTotal.add(amount);
	}
	
	public void increaseExpendituresTotal(BigDecimal amount) {
		expendituresTotal = expendituresTotal.add(amount);
	}
	
	public void increaseCommittedFundsTotal(BigDecimal amount) {
		committedFundsTotal = committedFundsTotal.add(amount);
	}
	
	public void increaseStipendTotal(BigDecimal amount) {
		stipendTotal = stipendTotal.add(amount);
	}
	
	public void increaseAwardsTotal(BigDecimal amount) {
		awardsTotal = awardsTotal.add(amount);
	}

	public BigDecimal getProposalsTotal() {
		return proposalsTotal;
	}

	public BigDecimal getExpendituresTotal() {
		return expendituresTotal;
	}

	public BigDecimal getCommittedFundsTotal() {
		return committedFundsTotal;
	}

	public BigDecimal getStipendTotal() {
		return stipendTotal;
	}

	public BigDecimal getAwardsTotal() {
		return awardsTotal;
	}
	
	public String toString() {
		return name + "\n\tProposals: " + proposalsTotal
					+ "\n\tExpenditures: " + expendituresTotal
					+ "\n\tCommitted Funds: " + committedFundsTotal
					+ "\n\tStipend: " + stipendTotal
					+ "\n\tAwards: " + awardsTotal;
	}
	
}
