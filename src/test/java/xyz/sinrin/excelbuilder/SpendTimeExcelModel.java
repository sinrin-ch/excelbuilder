package xyz.sinrin.excelbuilder;

/**
 * Created by sinrin on 2018/5/5.
 */
public class SpendTimeExcelModel
{
	@CellConfig(0)
	private Integer rKey;
	@CellConfig(1)
	private String departmentName;
	@CellConfig(2)
	private Integer two;
	@CellConfig(3)
	private Integer four;
	@CellConfig(4)
	private Integer six;
	@CellConfig(5)
	private Integer ten;
	@CellConfig(6)
	private Integer fourteen;
	@CellConfig(7)
	private Integer eighteen;

	public Integer getrKey()
	{
		return rKey;
	}

	public void setrKey(final Integer rKey)
	{
		this.rKey = rKey;
	}

	public String getDepartmentName()
	{
		return departmentName;
	}

	public void setDepartmentName(final String departmentName)
	{
		this.departmentName = departmentName;
	}

	public Integer getTwo()
	{
		return two;
	}

	public void setTwo(final Integer two)
	{
		this.two = two;
	}

	public Integer getFour()
	{
		return four;
	}

	public void setFour(final Integer four)
	{
		this.four = four;
	}

	public Integer getSix()
	{
		return six;
	}

	public void setSix(final Integer six)
	{
		this.six = six;
	}

	public Integer getTen()
	{
		return ten;
	}

	public void setTen(final Integer ten)
	{
		this.ten = ten;
	}

	public Integer getFourteen()
	{
		return fourteen;
	}

	public void setFourteen(final Integer fourteen)
	{
		this.fourteen = fourteen;
	}

	public Integer getEighteen()
	{
		return eighteen;
	}

	public void setEighteen(final Integer eighteen)
	{
		this.eighteen = eighteen;
	}

	public SpendTimeExcelModel()
	{
	}

	@Override
	public String toString()
	{
		return "SpendTimeExcelModel{" +
				       "rKey=" + rKey +
				       ", departmentName='" + departmentName + '\'' +
				       ", two=" + two +
				       ", four=" + four +
				       ", six=" + six +
				       ", ten=" + ten +
				       ", fourteen=" + fourteen +
				       ", eighteen=" + eighteen +
				       '}';
	}
}
