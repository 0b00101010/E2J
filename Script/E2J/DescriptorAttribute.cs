using System;
using E2J;

namespace E2J {
    public class DescriptorAttribute : Attribute {
        private string tableName;
        public string TableName => tableName;

        public DescriptorAttribute(string tableName) {
            this.tableName = tableName;
        }
    }   
}