package de.marcb.projects.exceltools;

import java.util.ArrayList;
import java.util.Collections;
import java.util.Iterator;
import java.util.List;
import java.util.Optional;
import java.util.stream.Stream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;

public class Column implements Iterable<Cell> {

	private final List<Cell> column;

	public Column() {
		column = new ArrayList<>();
	}

	public Column(final Cell cell) {
		column = new ArrayList<>();
		add(cell);
	}

	public Cell get(final int index) {
		return column.get(index);
	}

	public Optional<Object> getValue(final int index) {
		return Optional.ofNullable(getCellValue(column.get(index)));
	}

	public void add(final Cell cell) {
		column.add(cell);
	}

	public void addIfAbsent(final Cell cell) {
		if (!column.contains(cell)) {
			column.add(cell);
		}
	}

	public int count() {
		return column.size();
	}

	/**
	 * @see List#stream()
	 * @return
	 */
	public Stream<Cell> stream() {
		return column.stream();
	}

	@Override
	public Iterator<Cell> iterator() {
		return column.iterator();
	}

	public List<Cell> toList() {
		return Collections.unmodifiableList(column);
	}

	public boolean areOfCelltype(final CellType cellType) {
		for (Cell cell : this) {
			if (cell.getCellTypeEnum() != cellType) {
				return false;
			}
		}
		return true;
	}

	private Object getCellValue(final Cell cell) {
		final CellType type = cell.getCellTypeEnum();
		if (type == CellType.BOOLEAN) {
			return cell.getBooleanCellValue();
		}
		if (type == CellType.NUMERIC) {
			return cell.getNumericCellValue();
		}
		if (type == CellType.STRING || type == CellType.FORMULA) {
			return cell.getRichStringCellValue();
		}
		return null;
	}

	@Override
	public String toString() {
		StringBuilder builder = new StringBuilder("Column[");
		Iterator<Cell> iterator = iterator();
		while (iterator.hasNext()) {
			Cell cell = iterator.next();
			builder.append("Cell[" + cell.getCellTypeEnum().name() + " - " + getCellValue(cell) + "]");
			if (iterator.hasNext()) {
				builder.append(", ");
			}
		}
		builder.append("]");
		return builder.toString();
	}
}
