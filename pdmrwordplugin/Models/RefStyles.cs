using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace pdmrwordplugin.Models
{


    // NOTE: Generated code may require at least .NET Framework 4.5 or .NET Core/Standard 2.0.
    /// <remarks/>
    [System.SerializableAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(AnonymousType = true)]
    [System.Xml.Serialization.XmlRootAttribute(Namespace = "", IsNullable = false)]
    public partial class referencestyles
    {

        private referencestylesStyle[] styleField;

        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute("style")]
        public referencestylesStyle[] style
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

    /// <remarks/>
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

        private VolumeStyle volumestylefield;
        private IssueStyle issuestylefield;
        private PageStyle pagestylefield;
        private DateStyle datestylefield;

        /// <remarks/>
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

        /// <remarks/>
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

        public DateStyle date
        {
            get { return this.datestylefield; }
            set { this.datestylefield = value; }
        }

        public VolumeStyle volume
        {
            get { return this.volumestylefield; }
            set { this.volumestylefield = value; }
        }

        public IssueStyle issue
        {
            get { return this.issuestylefield; }
            set { this.issuestylefield = value; }
        }

        public PageStyle page
        {
            get { return this.pagestylefield; }
            set { this.pagestylefield = value; }
        }        

        /// <remarks/>
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

        /// <remarks/>
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

        /// <remarks/>
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

        private object useperiodField;

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
        public object useperiod
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


    [System.SerializableAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(AnonymousType = true)]
    public partial class VolumeStyle
    {
        private bool boldfield;
        private bool Italicfield;

        public bool bold
        {
            get { return this.boldfield; }
            set { this.boldfield = value; }
        }

        public bool italic
        {
            get { return this.Italicfield; }
            set { this.Italicfield = value; }
        }
    }

    public partial class IssueStyle
    {
        private bool boldfield;
        private bool Italicfield;

        public bool bold
        {
            get { return this.boldfield; }
            set { this.boldfield = value; }
        }

        public bool italic
        {
            get { return this.Italicfield; }
            set { this.Italicfield = value; }
        }
    }

    public partial class DateStyle
    {
        private bool boldfield;
        private bool Italicfield;

        public bool bold
        {
            get { return this.boldfield; }
            set { this.boldfield = value; }
        }

        public bool italic
        {
            get { return this.Italicfield; }
            set { this.Italicfield = value; }
        }
    }

    public partial class PageStyle
    {
        private bool ellifield;
        private string seprfield;
        private string removefield;

        public bool ellision
        {
            get { return this.ellifield; }
            set { this.ellifield = value; }
        }

        public string separator
        {
            get { return this.seprfield; }
            set { this.seprfield = value; }
        }

        public string omitchars
        {
            get { return this.removefield; }
            set { this.removefield = value; }
        }
    }


    /// <remarks/>
    [System.SerializableAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(AnonymousType = true)]
    public partial class referencestylesStyleSeparators
    {

        private string maxcountField;

        private string countField;

        private string etalField;

        private string authorField;

        private string twoauthorField;

        private string lastnameField;

        private string initialsField;

        private string lastauthorField;

        private string endField;

        private string beforeprefixField;

        /// <remarks/>
        public string maxcount
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
        public string count
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
