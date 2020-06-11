package src;

import java.awt.Component;
import java.util.List;

import javax.swing.JTable;
import javax.swing.JTextArea;
import javax.swing.border.EmptyBorder;
import javax.swing.table.TableCellRenderer;

public class Renderer extends JTextArea implements
TableCellRenderer {

    private static final long serialVersionUID = 1L;
    boolean                     limit               = false;

    public Renderer() {
        setLineWrap(true);
        setWrapStyleWord(true);
        setOpaque(true);
        setBorder(new EmptyBorder(-1, 2, -1, 2));
        this.limit = false;
        setRows(1);
    }

    public Renderer(int rows) {
        this();
        setRows(rows);
        this.limit = true;
    }

    @Override
    public Component getTableCellRendererComponent(JTable table, Object value,
            boolean isSelected, boolean hasFocus, int row, int column) {
        setText(value == null ? "" : value.toString());

        int dynamicHeight = getLineCount() > getRows() && this.limit ? getRows() : getLineCount();
        int newHeight = table.getRowHeight() * dynamicHeight;
        if (table.getRowHeight(row) != newHeight) {
            table.setRowHeight(row, newHeight);
        }

        if (isSelected) {
            setForeground(table.getSelectionForeground());
            setBackground(table.getSelectionBackground());
        }
        else {
            setForeground(table.getForeground());
            setBackground(table.getBackground());
        }

        return this;
    }
}