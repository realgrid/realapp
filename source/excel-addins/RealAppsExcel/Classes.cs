using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RealAppsExcel
{
    internal class ColumnHeader
    {
        public string text = "";
    }
    internal class ColumnEditor
    {
        public string type;
        public string datetimeFormat;
        public string editFormat;
        public string textAlignment;
        public Nullable<bool> showButtons = null;
    }
    internal class ColumnStyles
    {
        public string textAlignment;
        public string numberFormat;
    }

    internal class GridField
    {
        public string fieldName;
        public string dataType;
        public string datetimeFormat;
    }

    internal class GridColumn
    {
        public int width;
        public ColumnHeader header;

        public void setHeader(string text)
        {
            if (this.header == null)
            {
                this.header = new ColumnHeader();
            }
            header.text = text;
        }
    }

    internal class GridDataColumn : GridColumn
    {
        public string name;
        public string fieldName;
        public Nullable<bool> editable = null;
        public Nullable<bool> lookupDisplay = null;
        public string values;
        public string labels;
        public string valueSeperator;
        public ColumnEditor editor;
        public ColumnStyles styles;
        public void setEditor(string type)
        {
            if (this.editor == null)
            {
                this.editor = new ColumnEditor();
            }
            this.editor.type = type;
        }
        public void setStyles(string textAlignment = null, string numberFormat = null)
        {
            if (this.styles == null)
            {
                this.styles = new ColumnStyles();
            }
            if (textAlignment != null)
            {
                this.styles.textAlignment = textAlignment;
            }
            if (numberFormat != null)
            {
                this.styles.numberFormat = numberFormat;
            }
        }
    }

    internal class GridGroup : GridColumn
    {
        public string type = "group";
        public List<GridColumn> columns = new List<GridColumn>();
    }
}
