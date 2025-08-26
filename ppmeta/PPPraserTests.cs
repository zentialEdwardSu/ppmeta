using Xunit;
using System.Collections.Generic;

namespace ppmeta.Tests
{
    public class PPPraserTests
    {
        #region Basic Syntax Tests

        [Fact]
        public void Test_EmptyInput()
        {
            var result = PPParser.Parse("");
            Assert.False(result.HasErrors);
            Assert.Equal(0, result.Items.Count);
        }

        [Fact]
        public void Test_NullInput()
        {
            var result = PPParser.Parse(null);
            Assert.False(result.HasErrors);
            Assert.Equal(0, result.Items.Count);
        }

        [Fact]
        public void Test_SimpleFormatBlock()
        {
            var src = "[Title]\nHello World";
            var result = PPParser.Parse(src);
            Assert.False(result.HasErrors);
            Assert.Equal(1, result.Items.Count);
            Assert.Equal("Title", result.Items[0].Format);
            Assert.Equal("Hello World", result.Items[0].Content.Trim());
            Assert.Equal(0, result.Items[0].Placeholders.Count);
        }

        [Fact]
        public void Test_MultipleFormatBlocks()
        {
            var src = "[Title]\nContent1\n[Subtitle]\nContent2\n[Content]\nContent3";
            var result = PPParser.Parse(src);
            Assert.False(result.HasErrors);
            Assert.Equal(3, result.Items.Count);
            Assert.Equal("Title", result.Items[0].Format);
            Assert.Equal("Subtitle", result.Items[1].Format);
            Assert.Equal("Content", result.Items[2].Format);
        }

        [Fact]
        public void Test_EmptyFormatBlock()
        {
            var src = "[Title]\n[Subtitle]\nContent";
            var result = PPParser.Parse(src);
            Assert.False(result.HasErrors);
            Assert.Equal(2, result.Items.Count);
            Assert.Equal("", result.Items[0].Content.Trim());
            Assert.Equal("Content", result.Items[1].Content.Trim());
        }

        #endregion

        #region Basic PlaceHolder Parsing Tests

        [Fact]
        public void Test_CurrentScopePlaceholder()
        {
            var src = "[Title]\n$`name`=John\nHello";
            var result = PPParser.Parse(src);
            Assert.False(result.HasErrors);
            Assert.Equal(1, result.Items.Count);
            Assert.Equal("John", result.Items[0].Placeholders["name"]);
            Assert.Equal("Hello", result.Items[0].Content.Trim());
        }

        [Fact]
        public void Test_CurrentScopeAlternateSyntax()
        {
            var src = "[Title]\n$()`name`=John\nHello";
            var result = PPParser.Parse(src);
            Assert.False(result.HasErrors);
            Assert.Equal(1, result.Items.Count);
            Assert.Equal("John", result.Items[0].Placeholders["name"]);
        }

        [Fact]
        public void Test_GlobalPlaceholder()
        {
            var src = "$(g)`author`=Alice\n[Title]\nContent1\n[Subtitle]\nContent2";
            var result = PPParser.Parse(src);
            Assert.False(result.HasErrors);
            Assert.Equal(2, result.Items.Count);
            Assert.Equal("Alice", result.Items[0].Placeholders["author"]);
            Assert.Equal("Alice", result.Items[1].Placeholders["author"]);
        }

        [Fact]
        public void Test_GlobalPlaceholderUppercase()
        {
            var src = "$(G)`author`=Bob\n[Title]\nContent";
            var result = PPParser.Parse(src);
            Assert.False(result.HasErrors);
            Assert.Equal("Bob", result.Items[0].Placeholders["author"]);
        }

        [Fact]
        public void Test_TemporaryPlaceholder()
        {
            var src = "$(2)`temp`=TempValue\n[Block1]\nContent1\n[Block2]\nContent2\n[Block3]\nContent3";
            var result = PPParser.Parse(src);
            Assert.False(result.HasErrors);
            Assert.Equal(3, result.Items.Count);
            Assert.Equal("TempValue", result.Items[0].Placeholders["temp"]);
            Assert.Equal("TempValue", result.Items[1].Placeholders["temp"]);
            Assert.False(result.Items[2].Placeholders.ContainsKey("temp"));
        }

        #endregion

        #region Multi-line tests

        [Fact]
        public void Test_MultiLineValue()
        {
            var src = "[Title]\n$`desc`={\nLine 1\nLine 2\nLine 3\n}\nContent";
            var result = PPParser.Parse(src);
            Assert.False(result.HasErrors);
            Assert.Equal("Line 1\nLine 2\nLine 3", result.Items[0].Placeholders["desc"]);
        }

        [Fact]
        public void Test_EmptyMultiLineValue()
        {
            var src = "[Title]\n$`empty`={\n}\nContent";
            var result = PPParser.Parse(src);
            Assert.False(result.HasErrors);
            Assert.Equal("", result.Items[0].Placeholders["empty"]);
        }

        [Fact]
        public void Test_MultiLineValueWithEscaping()
        {
            var src = "[Title]\n$`desc`={\n\\[escaped\\]\n\\$dollar\n}\nContent";
            var result = PPParser.Parse(src);
            Assert.False(result.HasErrors);
            Assert.Equal("[escaped]\n$dollar", result.Items[0].Placeholders["desc"]);
        }

        [Fact]
        public void Test_GlobalMultiLineValue()
        {
            var src = "$(g)`global`={\nGlobal Line 1\nGlobal Line 2\n}\n[Title]\nContent";
            var result = PPParser.Parse(src);
            Assert.False(result.HasErrors);
            Assert.Equal("Global Line 1\nGlobal Line 2", result.Items[0].Placeholders["global"]);
        }

        #endregion

        #region Escaping Tests

        [Fact]
        public void Test_EscapedBrackets()
        {
            var src = "[Title]\n\\[Not a format\\]\nThis is content";
            var result = PPParser.Parse(src);
            Assert.False(result.HasErrors);
            Assert.Equal(1, result.Items.Count);
            Assert.Equal("[Not a format]\nThis is content", result.Items[0].Content.Trim());
        }

        [Fact]
        public void Test_EscapedDollar()
        {
            var src = "[Title]\n\\$`not a placeholder`=value\nContent";
            var result = PPParser.Parse(src);
            Assert.False(result.HasErrors);
            Assert.Equal("$`not a placeholder`=value\nContent", result.Items[0].Content.Trim());
        }

        [Fact]
        public void Test_EscapedFormatName()
        {
            var src = "[Title \\[with brackets\\]]\nContent";
            var result = PPParser.Parse(src);
            Assert.False(result.HasErrors);
            Assert.Equal("Title [with brackets]", result.Items[0].Format);
        }

        [Fact]
        public void Test_EscapedPlaceholderName()
        {
            var src = "[Title]\n$`name\\[bracket\\]`=value\nContent";
            var result = PPParser.Parse(src);
            Assert.False(result.HasErrors);
            Assert.Equal("value", result.Items[0].Placeholders["name[bracket]"]);
        }

        #endregion

        #region Variable Priority Tests

        [Fact]
        public void Test_VariablePriority_CurrentOverGlobal()
        {
            var src = "$(g)`name`=Global\n[Title]\n$`name`=Current\nContent";
            var result = PPParser.Parse(src);
            Assert.False(result.HasErrors);
            Assert.Equal("Current", result.Items[0].Placeholders["name"]);
        }

        [Fact]
        public void Test_VariablePriority_CurrentOverTemporary()
        {
            var src = "$(2)`name`=Temp\n[Title]\n$`name`=Current\nContent";
            var result = PPParser.Parse(src);
            Assert.False(result.HasErrors);
            Assert.Equal("Current", result.Items[0].Placeholders["name"]);
        }

        [Fact]
        public void Test_VariablePriority_TemporaryOverGlobal()
        {
            var src = "$(g)`name`=Global\n$(1)`name`=Temp\n[Title]\nContent";
            var result = PPParser.Parse(src);
            Assert.False(result.HasErrors);
            Assert.Equal("Temp", result.Items[0].Placeholders["name"]);
        }

        [Fact]
        public void Test_VariablePriority_Complex()
        {
            var src = "$(g)`name`=Global\n$(2)`name`=Temp\n[Block1]\n$`name`=Current1\nContent1\n[Block2]\nContent2\n[Block3]\nContent3";
            var result = PPParser.Parse(src);
            Assert.False(result.HasErrors);
            Assert.Equal(3, result.Items.Count);
            Assert.Equal("Current1", result.Items[0].Placeholders["name"]);
            Assert.Equal("Temp", result.Items[1].Placeholders["name"]);
            Assert.Equal("Global", result.Items[2].Placeholders["name"]);
        }

        #endregion

        #region Complex Syntax Combaination Tests

        [Fact]
        public void Test_ComplexSyntaxCombination()
        {
            var src = @"$(g)`author`=Global Author
$(g)`version`={
Version 1.0
Build 2024
}
$(3)`chapter`=Chapter One
[Title Slide]
$`title`=Main Title
$`subtitle`={
A comprehensive guide
to complex syntax
}
Welcome to the presentation

[Content Slide]
$`bullet1`=First point
$`bullet2`=Second point
Content here

[Summary]
$()`conclusion`={
Thank you for reading
Questions welcome
}
End of presentation

[Appendix]
Additional information";

            var result = PPParser.Parse(src);
            Assert.False(result.HasErrors);
            Assert.Equal(4, result.Items.Count);

            // Title Slide
            var titleSlide = result.Items[0];
            Assert.Equal("Title Slide", titleSlide.Format);
            Assert.Equal("Global Author", titleSlide.Placeholders["author"]);
            Assert.Equal("Version 1.0\nBuild 2024", titleSlide.Placeholders["version"]);
            Assert.Equal("Chapter One", titleSlide.Placeholders["chapter"]);
            Assert.Equal("Main Title", titleSlide.Placeholders["title"]);
            Assert.Equal("A comprehensive guide\nto complex syntax", titleSlide.Placeholders["subtitle"]);
            Assert.Equal("Welcome to the presentation", titleSlide.Content.Trim());

            // Content Slide
            var contentSlide = result.Items[1];
            Assert.Equal("Content Slide", contentSlide.Format);
            Assert.Equal("Global Author", contentSlide.Placeholders["author"]);
            Assert.Equal("Chapter One", contentSlide.Placeholders["chapter"]);
            Assert.Equal("First point", contentSlide.Placeholders["bullet1"]);
            Assert.Equal("Second point", contentSlide.Placeholders["bullet2"]);
            Assert.False(contentSlide.Placeholders.ContainsKey("title"));

            // Summary
            var summarySlide = result.Items[2];
            Assert.Equal("Summary", summarySlide.Format);
            Assert.Equal("Global Author", summarySlide.Placeholders["author"]);
            Assert.Equal("Chapter One", summarySlide.Placeholders["chapter"]);
            Assert.Equal("Thank you for reading\nQuestions welcome", summarySlide.Placeholders["conclusion"]);

            // Appendix (chapter expired)
            var appendixSlide = result.Items[3];
            Assert.Equal("Appendix", appendixSlide.Format);
            Assert.Equal("Global Author", appendixSlide.Placeholders["author"]);
            Assert.False(appendixSlide.Placeholders.ContainsKey("chapter"));
        }

        [Fact]
        public void Test_NestedEscaping()
        {
            var src = @"[Format \\[with\\] brackets]
$`complex\\$name`={
Line with \\[brackets\\]
Line with \\$dollar
}
Content with \\[escaped\\] and \\$escaped";

            var result = PPParser.Parse(src);
            Assert.False(result.HasErrors);
            Assert.Equal(1, result.Items.Count);
            Assert.Equal("Format \\[with\\] brackets", result.Items[0].Format);
            Assert.Equal("Line with [brackets]\nLine with $dollar", result.Items[0].Placeholders["complex$name"]);
            Assert.Equal("Content with [escaped] and $escaped", result.Items[0].Content.Trim());
        }

        #endregion

        #region Error Processing Tests

        [Fact]
        public void Test_Error_EmptyFormat()
        {
            var src = "[]\nContent";
            var result = PPParser.Parse(src);
            Assert.True(result.HasErrors);
            Assert.Contains("格式名称不能为空", result.Errors[0]);
        }

        [Fact]
        public void Test_Error_ShortFormat()
        {
            var src = "[]\nContent";
            var result = PPParser.Parse(src);
            Assert.True(result.HasErrors);
        }

        [Fact]
        public void Test_Error_PlaceholderOutsideFormat()
        {
            var src = "$`name`=value\nContent";
            var result = PPParser.Parse(src);
            Assert.True(result.HasErrors);
            Assert.Contains("当前作用域的变量必须在 [Format] 块内定义", result.Errors[0]);
        }

        [Fact]
        public void Test_Error_TempPlaceholderOutsideFormat()
        {
            var src = "$(2)`temp`=value\nContent";
            var result = PPParser.Parse(src);
            Assert.True(result.HasErrors);
            Assert.Contains("临时作用域的变量必须在 [Format] 块内定义", result.Errors[0]);
        }

        [Fact]
        public void Test_Error_ContentOutsideFormat()
        {
            var src = "Some content without format";
            var result = PPParser.Parse(src);
            Assert.True(result.HasErrors);
            Assert.Contains("内容必须在 [Format] 块内定义", result.Errors[0]);
        }

        [Fact]
        public void Test_Error_InvalidPlaceholderSyntax()
        {
            var src = "[Title]\n$invalid syntax\nContent";
            var result = PPParser.Parse(src);
            Assert.True(result.HasErrors);
            Assert.Contains("变量定义语法错误", result.Errors[0]);
        }

        [Fact]
        public void Test_Error_EmptyPlaceholderName()
        {
            var src = "[Title]\n$``=value\nContent";
            var result = PPParser.Parse(src);
            Assert.True(result.HasErrors);
            Assert.Contains("变量名不能为空", result.Errors[0]);
        }

        [Fact]
        public void Test_Error_InvalidTempScope()
        {
            var src = "[Title]\n$(0)`name`=value\nContent";
            var result = PPParser.Parse(src);
            Assert.True(result.HasErrors);
            Assert.Contains("临时作用域的块数必须大于0", result.Errors[0]);
        }

        [Fact]
        public void Test_Error_InvalidScopeDefinition()
        {
            var src = "[Title]\n$(invalid)`name`=value\nContent";
            var result = PPParser.Parse(src);
            Assert.True(result.HasErrors);
            Assert.Contains("无效的作用域定义", result.Errors[0]);
        }

        [Fact]
        public void Test_Error_UnterminatedMultiLine()
        {
            var src = "[Title]\n$`name`={\nLine 1\nLine 2";
            var result = PPParser.Parse(src);
            Assert.True(result.HasErrors);
            Assert.Contains("多行值缺少结束符", result.Errors[0]);
        }

        [Fact]
        public void Test_Error_MultipleErrors()
        {
            var src = "Content without format\n$(0)`invalid`=value\n[]\n$`empty`={\nunterminated";
            var result = PPParser.Parse(src);
            Assert.True(result.HasErrors);
            Assert.True(result.Errors.Count >= 3);
        }

        #endregion

        #region Edge case Tests

        [Fact]
        public void Test_WhitespaceHandling()
        {
            var src = "  [  Title  ]  \n  $`name`  =  value  \n  Content  ";
            var result = PPParser.Parse(src);
            Assert.False(result.HasErrors);
            Assert.Equal("Title", result.Items[0].Format);
            Assert.Equal("value", result.Items[0].Placeholders["name"]);
        }

        [Fact]
        public void Test_EmptyLines()
        {
            var src = "\n\n[Title]\n\n$`name`=value\n\nContent\n\n";
            var result = PPParser.Parse(src);
            Assert.False(result.HasErrors);
            Assert.Equal(1, result.Items.Count);
        }

        [Fact]
        public void Test_SpecialCharactersInValues()
        {
            var src = "[Title]\n$`special`=!@#$%^&*()_+-={}[]|\\:;\"'<>,.?/~`\nContent";
            var result = PPParser.Parse(src);
            Assert.False(result.HasErrors);
            Assert.Equal("!@#$%^&*()_+-={}[]|\\:;\"'<>,.?/~`", result.Items[0].Placeholders["special"]);
        }

        [Fact]
        public void Test_UnicodeCharacters()
        {
            var src = "[标题]\n$`名称`=值\n内容";
            var result = PPParser.Parse(src);
            Assert.False(result.HasErrors);
            Assert.Equal("标题", result.Items[0].Format);
            Assert.Equal("值", result.Items[0].Placeholders["名称"]);
            Assert.Equal("内容", result.Items[0].Content.Trim());
        }

        #endregion
    }
}
