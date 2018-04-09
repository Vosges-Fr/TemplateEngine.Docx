using System;

namespace TemplateEngine.Docx
{
	[ContentItemName("Field")]
	public class FieldContent : HiddenContent<FieldContent>, IEquatable<FieldContent>
	{
        public FieldContent()
        {
            
        }

        /// <summary>
        /// Generic form of field content.
        /// Not to use with checkboxes.
        /// </summary>
        /// <param name="name"></param>
        /// <param name="value"></param>
        public FieldContent(string name, string value)
        {
            Name = name;
            Value = value;
        }

        /// <summary>
        /// Specifically for checkbox fields.
        /// </summary>
        /// <param name="name"></param>
        /// <param name="value"></param>
        public FieldContent(string name, bool value) : this(name, null)
        {
            BoolValue = value;
        }


        /// <summary>
        /// Specifically for checkbox fields.
        /// </summary>
        /// <param name="name"></param>
        /// <param name="value"></param>
        public FieldContent(string name, DateTime value) : this(name, null)
        {
            DateTimeValue = value;
        }

        public string Value { get; set; }

        private bool _BoolValue;
        public bool BoolValue {
            get { return _BoolValue; }
            set { _BoolValue = value; Value = value.ToString(); }
        }

        private DateTime _DateTimeValue;
        public DateTime DateTimeValue
        {
            get { return _DateTimeValue; }
            set { _DateTimeValue = value; Value = value.ToBinary().ToString(); }
        }

        #region Equals

        public bool Equals(FieldContent other)
		{
			if (other == null) return false;

			return Name.Equals(other.Name) &&
			       Value.Equals(other.Value);
		}

		public override bool Equals(IContentItem other)
		{
			if (!(other is FieldContent)) return false;

			return Equals((FieldContent)other);
		}

		public override int GetHashCode()
		{
			return new { Name, Value }.GetHashCode();
		}

        #endregion
    }
}
