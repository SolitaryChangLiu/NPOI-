using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Panduo.NPOIDemo
{
    public class ExcelHeadView
    {
        /// <summary>
        /// 列名
        /// </summary>
        public string CellName { get; set; }
        /// <summary>
        /// 父列表头下是否是第一个子列
        /// </summary>
        public bool IsFirstSonCell { get; set; }
        /// <summary>
        /// 父列表头下子列的列数
        /// </summary>
        public int Span { get; set; }
        /// <summary>
        /// 父列表头的名称
        /// </summary>
        public string FatherCellName { get; set; }
        /// <summary>
        /// 表名称
        /// </summary>
        public string TableName { get; set; }
        /// <summary>
        /// 该列在表的索引位置
        /// </summary>
        public int index { get; set; }

    }
}