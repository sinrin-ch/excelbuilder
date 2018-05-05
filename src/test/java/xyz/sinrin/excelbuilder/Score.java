package xyz.sinrin.excelbuilder;

import java.time.LocalDate;
import java.util.Date;

public class Score {
    @CellConfig(0)
    private String name;
    @CellConfig(1)
    private Integer age;
    @CellConfig(2)
    private Double height;
    @CellConfig(3)
    private LocalDate birthday;

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public Integer getAge() {
        return age;
    }

    public void setAge(Integer age) {
        this.age = age;
    }

    public Double getHeight() {
        return height;
    }

    public void setHeight(Double height) {
        this.height = height;
    }

    public LocalDate getBirthday() {
        return birthday;
    }

    public void setBirthday(LocalDate birthday) {
        this.birthday = birthday;
    }

    public Score() {
    }

    @Override
    public String toString() {
        return "Score{" +
                "name='" + name + '\'' +
                ", age=" + age +
                ", height=" + height +
                ", birthday=" + birthday +
                '}';
    }
}
