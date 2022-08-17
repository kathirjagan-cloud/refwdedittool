using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace pdmrwordplugin.Models
{
    [System.SerializableAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(AnonymousType = true)]
    [System.Xml.Serialization.XmlRootAttribute(Namespace = "", IsNullable = false)]
    public partial class referencestyles
    {
        private referencestylesStyle styleField;
        public referencestylesStyle style
        {
            get
            {
                return this.styleField;
            }
            set
            {
                this.styleField = value;
            }
        }
    }

    [System.SerializableAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(AnonymousType = true)]
    public partial class referencestylesStyle
    {

        private string patternField;

        private referencestylesStyleJournal journalField;

        private string authorpatternField;

        private referencestylesStyleSeparators separatorsField;

        private string nameField;
        public string pattern
        {
            get
            {
                return this.patternField;
            }
            set
            {
                this.patternField = value;
            }
        }

        public referencestylesStyleJournal journal
        {
            get
            {
                return this.journalField;
            }
            set
            {
                this.journalField = value;
            }
        }


        public string authorpattern
        {
            get
            {
                return this.authorpatternField;
            }
            set
            {
                this.authorpatternField = value;
            }
        }

        public referencestylesStyleSeparators separators
        {
            get
            {
                return this.separatorsField;
            }
            set
            {
                this.separatorsField = value;
            }
        }

        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string name
        {
            get
            {
                return this.nameField;
            }
            set
            {
                this.nameField = value;
            }
        }
    }

    /// <remarks/>
    [System.SerializableAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(AnonymousType = true)]
    public partial class referencestylesStyleJournal
    {

        private bool abbreviationField;

        private bool useperiodField;

        private bool italicField;

        /// <remarks/>
        public bool abbreviation
        {
            get
            {
                return this.abbreviationField;
            }
            set
            {
                this.abbreviationField = value;
            }
        }

        /// <remarks/>
        public bool useperiod
        {
            get
            {
                return this.useperiodField;
            }
            set
            {
                this.useperiodField = value;
            }
        }

        /// <remarks/>
        public bool italic
        {
            get
            {
                return this.italicField;
            }
            set
            {
                this.italicField = value;
            }
        }
    }

    /// <remarks/>
    [System.SerializableAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(AnonymousType = true)]
    public partial class referencestylesStyleSeparators
    {

        private int maxcountField;

        private int countField;

        private string etalField;

        private string authorField;

        private string twoauthorField;

        private string lastnameField;

        private string initialsField;

        private string lastauthorField;

        private string endField;

        private string beforeprefixField;

        /// <remarks/>
        public int maxcount
        {
            get
            {
                return this.maxcountField;
            }
            set
            {
                this.maxcountField = value;
            }
        }

        /// <remarks/>
        public int count
        {
            get
            {
                return this.countField;
            }
            set
            {
                this.countField = value;
            }
        }

        /// <remarks/>
        public string etal
        {
            get
            {
                return this.etalField;
            }
            set
            {
                this.etalField = value;
            }
        }

        /// <remarks/>
        public string author
        {
            get
            {
                return this.authorField;
            }
            set
            {
                this.authorField = value;
            }
        }

        /// <remarks/>
        public string twoauthor
        {
            get
            {
                return this.twoauthorField;
            }
            set
            {
                this.twoauthorField = value;
            }
        }

        /// <remarks/>
        public string lastname
        {
            get
            {
                return this.lastnameField;
            }
            set
            {
                this.lastnameField = value;
            }
        }

        /// <remarks/>
        public string initials
        {
            get
            {
                return this.initialsField;
            }
            set
            {
                this.initialsField = value;
            }
        }

        /// <remarks/>
        public string lastauthor
        {
            get
            {
                return this.lastauthorField;
            }
            set
            {
                this.lastauthorField = value;
            }
        }

        /// <remarks/>
        public string end
        {
            get
            {
                return this.endField;
            }
            set
            {
                this.endField = value;
            }
        }

        /// <remarks/>
        public string beforeprefix
        {
            get
            {
                return this.beforeprefixField;
            }
            set
            {
                this.beforeprefixField = value;
            }
        }
    }
    internal class RefStyles
    {

    }
}
