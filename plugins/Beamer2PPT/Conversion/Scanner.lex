%x str, overlay, optional, pre_overlay, pre_optional, tabular_arg, boverlay, boptional, bpre_overlay, bpre_optional, math, comment, verb
%s body

%using SimpleConverter.Contract;

%namespace SimpleConverter.Plugin.Beamer2PPT

ws          [ \t]*
wsp         [ \t]+
wsl         {ws}\r?\n?{ws}
envBegin    \\begin{wsl}
envEnd      \\end{wsl}

%{
    // global variables
    bool tabular = false;   // info if we are in tabular env., inclusive states don't seem to work and exclusive are out of question
    bool inBody = false;    // info if we are in body of document (needed for processing white space characters)
    char verbBoundary = '\0';  // boundary character for \verb command

    public String Filename { get; set; }
%}

%%
%{
    // local variables
    string unformattedText = "";
    int spaces = 0;
    int nls = 0;
    int tbl = 0;
%}

// Parsing of LaTeX macros
// =============================================================================

// Preambule
// -----------------------------------------------------------------------------
\\documentclass { BEGIN(pre_optional); return (int) Tokens.DOCUMENTCLASS; }
\\usepackage    { BEGIN(pre_optional); return (int) Tokens.USEPACKAGE; }
\\title         { return (int) Tokens.TITLE; }
\\author        { BEGIN(pre_optional); return (int) Tokens.AUTHOR; }
\\institute     { BEGIN(pre_optional); return (int) Tokens.INSTITUTE; }
\\today         { return (int) Tokens.TODAY; }
\\date          { return (int) Tokens.DATE; }

// Environments ( + environments specific)
// -----------------------------------------------------------------------------
{envBegin}\{document\}      { inBody = true; BEGIN(body); return (int) Tokens.BEGIN_DOCUMENT; }
{envEnd}\{document\}        { inBody = false; BEGIN(INITIAL); return (int) Tokens.END_DOCUMENT; }
{envBegin}\{frame\}         { BEGIN(pre_optional); return (int) Tokens.BEGIN_FRAME; }
{envEnd}\{frame\}           { return (int) Tokens.END_FRAME; }
{envBegin}\{itemize\}       { BEGIN(pre_optional); return (int) Tokens.BEGIN_ITEMIZE; }
{envEnd}\{itemize\}         { return (int) Tokens.END_ITEMIZE; }
{envBegin}\{enumerate\}     { BEGIN(pre_optional); return (int) Tokens.BEGIN_ENUMERATE; }
{envEnd}\{enumerate\}       { return (int) Tokens.END_ENUMERATE; }
{envBegin}\{description\}   { BEGIN(pre_optional); return (int) Tokens.BEGIN_DESCRIPTION; }
{envEnd}\{description\}     { return (int) Tokens.END_DESCRIPTION; }
{envBegin}\{tabular\}{ws}(\[[bct]\]{ws})?\{     { tabular = true; BEGIN(tabular_arg); tbl = 0; unformattedText = ""; return (int) Tokens.BEGIN_TABULAR; }
{envEnd}\{tabular\}         { tabular = false; return (int) Tokens.END_TABULAR; }
\\item                      { BEGIN(pre_overlay); return (int) Tokens.ITEM; }
\\multicolumn               { return (int) Tokens.MULTICOLUMN; }

// Beamer specific
// -----------------------------------------------------------------------------
\\frame         { BEGIN(pre_optional); return (int) Tokens.FRAME; }
\\frametitle    { BEGIN(pre_overlay); return (int) Tokens.FRAMETITLE; }
\\framesubtitle { BEGIN(pre_overlay); return (int) Tokens.FRAMESUBTITLE; }
\\pause         { return (int) Tokens.PAUSE; }
\\usetheme      { return (int) Tokens.USETHEME; }

// Text formatting
// -----------------------------------------------------------------------------
\\textbf        { BEGIN(pre_overlay); return (int) Tokens.TEXTBF; }
\\texttt        { BEGIN(pre_overlay); return (int) Tokens.TEXTTT; }
\\textit        { BEGIN(pre_overlay); return (int) Tokens.TEXTIT; }
\\textsc        { BEGIN(pre_overlay); return (int) Tokens.TEXTSC; }
\\bfseries[ ]?      { return (int) Tokens.BFSERIES; }
\\ttfamily[ ]?      { return (int) Tokens.TTFAMILY; }
\\itshape[ ]?       { return (int) Tokens.ITSHAPE; }
\\scshape[ ]?       { return (int) Tokens.SCSHAPE; }
\\tiny[ ]?          { return (int) Tokens.TINY; }
\\scriptsize[ ]?    { return (int) Tokens.SCRIPTSIZE; }
\\footnotesize[ ]?  { return (int) Tokens.FOOTNOTESIZE; }
\\small[ ]?         { return (int) Tokens.SMALL; }
\\normalsize[ ]?    { return (int) Tokens.NORMALSIZE; }
\\large[ ]?         { return (int) Tokens.LARGE; }
\\Large[ ]?         { return (int) Tokens.LARGE2; }
\\LARGE[ ]?         { return (int) Tokens.LARGE3; }
\\huge[ ]?          { return (int) Tokens.HUGE; }
\\Huge[ ]?          { return (int) Tokens.HUGE2; }
\\color         { BEGIN(pre_overlay); return (int) Tokens.COLOR; }
\\textcolor     { BEGIN(pre_overlay); return (int) Tokens.TEXTCOLOR; }
\\underline     { BEGIN(pre_overlay); return (int) Tokens.UNDERLINE; }
\\and           { return (int) Tokens.AND; }

// Images
// -----------------------------------------------------------------------------
\\includegraphics   { BEGIN(pre_optional); return (int) Tokens.INCLUDEGRAPHICS; }
\\graphicspath      { return (int) Tokens.GRAPHICSPATH; }

// Other
// -----------------------------------------------------------------------------
\\titlepage     { return (int) Tokens.TITLEPAGE; }
\\hline         { return (int) Tokens.HLINE; }
\\cline         { return (int) Tokens.CLINE; }
\\section       { return (int) Tokens.SECTION; }
\\subsection    { return (int) Tokens.SUBSECTION; }
\\subsubsection { return (int) Tokens.SUBSUBSECTION; }
\\LaTeX[ ]?     { BEGIN(str); unformattedText = @"LaTeX"; spaces = 0; nls = 0; }
// New paragraph
\\\\|\\cr       { BEGIN(pre_optional); if(tabular) return (int) Tokens.ENDROW; else return (int) Tokens.NL; }
\\newline       { return (int) Tokens.NL; }

\\#                       BEGIN(str); unformattedText += @"#"; spaces = 0; nls = 0;
\\\$                      BEGIN(str); unformattedText += @"$"; spaces = 0; nls = 0;
\\%                       BEGIN(str); unformattedText += @"%"; spaces = 0; nls = 0;
\\textasciicircum\{\}     BEGIN(str); unformattedText += @"^"; spaces = 0; nls = 0;
\\&                       BEGIN(str); unformattedText += @"&"; spaces = 0; nls = 0;
\\_                       BEGIN(str); unformattedText += @"_"; spaces = 0; nls = 0;
\\\{                      BEGIN(str); unformattedText += @"{"; spaces = 0; nls = 0;
\\\}                      BEGIN(str); unformattedText += @"}"; spaces = 0; nls = 0;
\\~\{\}                   BEGIN(str); unformattedText += @"~"; spaces = 0; nls = 0;
\\textbackslash\{\}       BEGIN(str); unformattedText += @"\"; spaces = 0; nls = 0;
\\textpipe(\{\})?         BEGIN(str); unformattedText += @"|"; spaces = 0; nls = 0;

// Short space, todo copy to plain text loop
\\[ ]           { BEGIN(str); unformattedText += @" "; spaces = 1; nls = 0; }
\{              { return '{'; }
\}              { return '}'; }
&               { return '&'; }
// Comments
%.*\r?\n?{ws}    {/*ignore*/}

// White space at the begining of line -> really needed?
^{wsp}         {/*ignore*/}

\$              BEGIN(math); unformattedText = "$"; spaces = 0; nls = 0;

// Verbatim
// -----------------------------------------------------------------------------
\\verb.        { BEGIN(verb); unformattedText = ""; spaces = 0; nls = 0; verbBoundary = yytext[yytext.Length - 1]; }
<verb> {
    .          {
                    if(yytext[0] != verbBoundary)
                        unformattedText += yytext;
                    else {
                        if(inBody)
                            BEGIN(body);
                        else
                            BEGIN(INITIAL);

                        yylval.Text = unformattedText;
                        return (int) Tokens.VERB;
                    }
               }
}

// Math (when math is implemented change token type to MATH)
// -----------------------------------------------------------------------------
<math> {
        [^\$]   unformattedText += yytext;
        \$      if(inBody) BEGIN(body); else BEGIN(INITIAL); yylval.Text = unformattedText + "$"; return (int) Tokens.STRING;
    }

// Overlay specification
// -----------------------------------------------------------------------------
<pre_overlay> {
        {wsl}<                  BEGIN(overlay); unformattedText = "";
        {wsl}[^<]               BEGIN(pre_optional); yyless(0);
    }
<overlay> {
        [^>]*                   unformattedText += yytext;
        >                       BEGIN(pre_optional); yylval.Text = unformattedText; return (int) Tokens.OVERLAY;
    }

// Optional parameters
// -----------------------------------------------------------------------------
<pre_optional> {
        {wsl}\[                 BEGIN(optional); unformattedText = "";
        {wsl}[^[]               if(inBody) BEGIN(body); else BEGIN(INITIAL); yyless(0);
    }
<optional> {
        [^\]]*                  unformattedText += yytext;
        \]                      if(inBody) BEGIN(body); else BEGIN(INITIAL); yylval.Text = unformattedText; return (int) Tokens.OPTIONAL;
    }

// Tabular settings
// -----------------------------------------------------------------------------
<tabular_arg> {
        \{                      tbl++; unformattedText += @"{";
        \}                      {
                                    tbl--; 
                                    if(tbl < 0) {
                                        yylval.Text = unformattedText;
                                        BEGIN(INITIAL);
                                        return (int) Tokens.STRING;
                                    }
                                    unformattedText += @"}";
                                }
        [^{}]*                  unformattedText += yytext;
}

// Plain text
// -----------------------------------------------------------------------------
<body> [^#\$%\^&_\{\}\\]    BEGIN(str); yyless(0); unformattedText = ""; spaces = 0; nls = 0;

[^#\$%\^&_\{\}\\[:IsWhiteSpace:]]    BEGIN(str); yyless(0); unformattedText = ""; spaces = 0; nls = 0;

<str> {
        [^#\$%\^&_\{\}~\\ \n\t\r]*      { unformattedText += yytext; spaces = 0; nls = 0; }
        " "|\t                     {    // eat multiple spaces
                                        spaces++;
                                        if(spaces == 1)
                                            unformattedText += @" ";
                                   }
        \r                         { /* ignore */ }
        ~                          {    spaces++; unformattedText += @" "; }
        \n                         {    // eat multiple new lines
                                        nls++;
                                        if(nls == 1 && spaces == 0) {   // if one empty line add space
                                            unformattedText += @" ";
                                            spaces++;
                                        }
                                        if(nls == 2) {  // if two empty lines remove space and add new line
                                            // need to remove last space
                                            if(unformattedText.Length > 0 && unformattedText.EndsWith(" "))
                                                unformattedText.Remove(unformattedText.Length - 1, 1);
                                            unformattedText += "\n";
                                        }
                                   }
        // escape sequences
        \\#                        unformattedText += @"#"; spaces = 0; nls = 0;
        \\\$                       unformattedText += @"$"; spaces = 0; nls = 0;
        \\%                        unformattedText += @"%"; spaces = 0; nls = 0;
        \\textasciicircum\{\}      unformattedText += @"^"; spaces = 0; nls = 0;
        \\&                        unformattedText += @"&"; spaces = 0; nls = 0;
        \\_                        unformattedText += @"_"; spaces = 0; nls = 0;
        \\\{                       unformattedText += @"{"; spaces = 0; nls = 0;
        \\\}                       unformattedText += @"}"; spaces = 0; nls = 0;
        \\~\{\}                    unformattedText += @"~"; spaces = 0; nls = 0;
        \\textbackslash\{\}        unformattedText += @"\"; spaces = 0; nls = 0;
        \\textpipe(\{\})?          unformattedText += @"|"; spaces = 0; nls = 0;
        \\LaTeX[ ]?                unformattedText += @"LaTeX"; spaces = 0; nls = 0; // todo: fix possible parameters
        %.*\n?{ws}                 {/* ignore comment inside plaintext */}
        // end of plain text
        [#\$\^&_\{\}\\]            {
                                        if(inBody)
                                            BEGIN(body);
                                        else
                                            BEGIN(INITIAL);

                                        yyless(0);

                                        // return text only if it contains text or new line
                                        if(unformattedText.IndexOf('\n') != -1 || unformattedText.Trim().Length != 0) 
                                        {
                                            yylval.Text = unformattedText;
                                            return (int) Tokens.STRING;
                                        }
                                   }
    }

// Block comment (comment environment)
// ----------------------------------------------------------------------------
{envBegin}\{comment\}           BEGIN(comment);

<comment> {
        [^\\]*                  {/* ignore */}
        \\                      {/* ignore */}
        {envEnd}\{comment\}     if(inBody) BEGIN(body); else BEGIN(INITIAL);
    }

// Unknown commands etc.
// ============================================================================
{envBegin}\{[^\}]+\}   { BEGIN(bpre_overlay); printWarning("Unknown environment " + yytext); }
{envEnd}\{[^\}]+\}     { BEGIN(bpre_overlay); printWarning("Unknown environment " + yytext); }

\\[[:IsLetter:]]+      { BEGIN(bpre_overlay); printWarning("Unknown command " + yytext); }

// Unknown command overlay specification - eat and ignore
// ----------------------------------------------------------------------------
<bpre_overlay> {
        {wsl}<                  BEGIN(boverlay);
        {wsl}[^<]               BEGIN(bpre_optional); yyless(0);
    }
<boverlay> {
        [^>]*                   {}
        >                       BEGIN(bpre_optional);
    }

//  Unknown command optional parameters - eat and ignore
// ----------------------------------------------------------------------------
<bpre_optional> {
        {wsl}\[                 BEGIN(boptional);
        {wsl}[^[]               if(inBody) BEGIN(body); else BEGIN(INITIAL); yyless(0);
    }
<boptional> {
        [^\]]*                  {}
        \]                      if(inBody) BEGIN(body); else BEGIN(INITIAL);
    }

<<EOF>>                            {/* to process, or not to process? */}

/* ------------------------------------------ */
%{
	yylloc = new QUT.Gppg.LexLocation(tokLin, tokCol, tokELin, tokECol);
%}
/* ------------------------------------------ */

%%

override public void yyerror(string format, params object[] args)
{
    string tmp;

    tmp = System.String.Format("{0}:{1}:{2} - ", Filename, yyline, yycol);
    Messenger.Instance.SendMessage(tmp + format, SimpleConverter.Contract.MessageLevel.ERROR);
}

/// <summary>
/// Print warning message
/// </summary>
/// <param name="message">Message</param>
private void printWarning(string message)
{
    string tmp;

    tmp = System.String.Format("{0}:{1}:{2} - ", Filename, yyline, yycol);
    Messenger.Instance.SendMessage(tmp + message, SimpleConverter.Contract.MessageLevel.WARNING);
}
