using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;
using System.IO;
using System.Xml.Xsl;
using System.Xml.XPath;
using System.Xml;
using System.Runtime.CompilerServices;

namespace NUnit2ReportTool
{
	/// <summary>
	/// 武汉平台开发部.测试:周睿
	/// 时间：2014-05-12；
	/// 扩展了合并多个单元测试用例的功能，
	/// </summary>
	class Program
	{
		static readonly string s_BasePath = AppDomain.CurrentDomain.BaseDirectory;
		static void Main(string[] args)
		{

			//XmlFileSet是nunit工具执行后生成的结果XML
			string xmlFileSetPath = Path.Combine(s_BasePath, "XmlFileSet");
			//注释文件,不过我这里注释一直都出不来,
			string xmlSummariesPath = Path.Combine(s_BasePath, "XmlSummaries");

		
			if( Directory.Exists(xmlFileSetPath) == false ) {
				throw new DirectoryNotFoundException("路径不存在:" + xmlFileSetPath+"需要目录：XmlFileSet");
			}

			if( Directory.Exists(xmlSummariesPath) == false ) {
				throw new DirectoryNotFoundException("路径不存在:" + xmlSummariesPath + "需要目录：XmlSummaries");
			}

			//获得目录下面所有的XML
			List<string> pathList = Directory.GetFiles(xmlFileSetPath, "*.xml", SearchOption.AllDirectories).ToList<string>();

			if(pathList.Count()==0 )
				throw new FileNotFoundException("文件不存在" );


			//有多份XML配置的,需要做合并的动作
			if( pathList.Count >=2 ) {
				string tempPath = CombineDocument(pathList);
				pathList.Clear();
				pathList.Add(tempPath);

			}

			List<string> summariesPathList = Directory.GetFiles(xmlSummariesPath, "*.xml", SearchOption.AllDirectories).ToList<string>();
			if( xmlSummariesPath.Count() == 0 )
				throw new FileNotFoundException("文件不存在");


			//开始执行
			NUnit2ReportTask report = new NUnit2ReportTask();
			report.Todir = Path.Combine(s_BasePath, "NunitReport_" + DateTime.Now.ToString("yyyyMMdd"));
			report.OutFilename = string.Format("{0}.html", "单元测试报告");
			report.XmlFileSet = pathList;
			report.XmlSummaries = summariesPathList;
			//report.Format = "frames";
			report.InitializeTask();
			report.ExecuteTask();
			System.Diagnostics.Process.Start(Path.Combine(report.Todir, report.OutFilename));
		}


		#region 合并XML

		[MethodImpl(MethodImplOptions.Synchronized)]
		static string CombineDocument(List<string> pathList)
		{
			XmlDocument document = LoadXml(pathList[0]);
			XmlNode firstNode = document.SelectSingleNode("/test-results/test-suite[last()]");
			for( int i = 1; i < pathList.Count; i++ ) {

				if( Path.GetFileName(pathList[i]) != "合并之后的XML.xml" ) {
					//加载第一个XML
					XmlNode node = LoadXml(pathList[i]).SelectSingleNode("/test-results/test-suite");

					firstNode.CreateNavigator().InsertAfter(node.CreateNavigator());
				}
			}
			//修改用例数目
			int count = document.SelectNodes("//test-case").Count;
			(document.SelectSingleNode("/test-results[@total]") as XmlElement).SetAttribute("total", count.ToString());

			File.WriteAllText(Path.Combine(s_BasePath, "合并之后的XML.xml"), document.OuterXml);
			return Path.Combine(s_BasePath, "合并之后的XML.xml");
		}
	    static XmlDocument LoadXml(string path)
		{
			XmlDocument document = new XmlDocument();
			document.Load(path);
			return document;
		}
		#endregion
	}
	#region 扩展报表 BY ZR
	public class NUnit2ReportTask
	{
		private const string XSL_DEF_FILE_NOFRAME = "NUnit-NoFrame.xsl";
		private const string XSL_DEF_FILE_FRAME = "NUnit-Frame.xsl";
		private const string XSL_DEF_FILE_SUMMARY = "NReport-Summary.xsl";//"NReport-Summary.xsl";

		private const string XSL_I18N_FILE = "i18n.xsl";

		private string _outFilename = "index.htm";
		private string _todir = "";

		private List<string> _fileset = new List<string>();
		private XmlDocument _FileSetSummary;

		private List<string> _summaries = new List<string>();
		private string _tempXmlFileSummarie = "";

		private string _xslFile = "";
		private string _i18nXsl = "";
		private string _summaryXsl = "";
		private string _openDescription = "yes";
		private string _language = "chs";
		private string _format = "noframes";
		private string _nantLocation = AppDomain.CurrentDomain.BaseDirectory;
		private XsltArgumentList _xsltArgs;

		/// <summary>
		/// The format of the generated report. 
		/// Must be "noframes" or "frames". 
		/// Default to "noframes".
		/// </summary>
		public string Format
		{
			get { return _format; }
			set { _format = value; }
		}

		/// <summary>
		/// The output language.
		/// </summary>
		public string Language
		{
			get { return _language; }
			set { _language = value; }
		}

		/// <summary>
		/// Open all description method. Default to "false".
		/// </summary>
		public string OpenDescription
		{
			get { return _openDescription; }
			set { _openDescription = value; }
		}

		/// <summary>
		/// The directory where the files resulting from the transformation should be written to.
		/// </summary>
		public string Todir
		{
			get { return _todir; }
			set { _todir = value; }
		}

		/// <summary>
		/// Index of the Output HTML file(s).
		/// Default to "index.htm".
		/// </summary>
		public string OutFilename
		{
			get { return _outFilename; }
			set { _outFilename = value; }
		}


		/// <summary>
		/// Set of XML files to use as input
		/// </summary>
		public List<string> XmlFileSet
		{
			get { return _fileset; }
			set { _fileset = value; }
		}

		/// <summary>
		/// Set of XML files to use as input
		/// </summary>
		public List<string> XmlSummaries
		{
			get { return _summaries; }
			set { _summaries = value; }
		}


		///<summary>
		///Initializes task and ensures the supplied attributes are valid.
		///</summary>
		///<param name="taskNode">Xml node used to define this task instance.</param>
		public void InitializeTask()
		{
			Assembly thisAssm = Assembly.GetExecutingAssembly();

#if ECHO_MODE
				Console.WriteLine ("Location : "+thisAssm.CodeBase);
#endif

			_nantLocation = Path.GetDirectoryName(thisAssm.CodeBase);//(thisAssm.Location


			if( this.Format == "noframes" ) {
				_xslFile = Path.Combine(_nantLocation, XSL_DEF_FILE_NOFRAME);
			}

			_i18nXsl = Path.Combine(_nantLocation, XSL_I18N_FILE);
			_summaryXsl = Path.Combine(_nantLocation, XSL_DEF_FILE_SUMMARY);

			if( this.XmlFileSet.Count == 0 ) {
				throw new Exception("NUnitReport fileset cannot be empty!");
			}

			foreach( string file in this.XmlSummaries ) {
				_tempXmlFileSummarie = file;
			}

			// Get the Nant, OS parameters
			_xsltArgs = GetPropertyList();

			//Create directory if ...
			if( this.Todir != "" ) {
				Directory.CreateDirectory(this.Todir);
			}

		}

		/// <summary>
		/// This is where the work is done
		/// </summary>
		public void ExecuteTask()
		{
			_FileSetSummary = CreateSummaryXmlDoc();

			foreach( string file in this.XmlFileSet ) {
				XmlDocument source = new XmlDocument();
				source.Load(file);
				XmlNode node = _FileSetSummary.ImportNode(source.DocumentElement, true);
				_FileSetSummary.DocumentElement.AppendChild(node);
			}

			//
			// prepare properties and transform
			//
			try {
				if( this.Format == "noframes" ) {

					XslTransform xslTransform = new XslTransform();
					xslTransform.Load(_xslFile);

					// xmlReader hold the first transformation
					using (XmlReader xmlReader = xslTransform.Transform(_FileSetSummary, _xsltArgs)) {
						
						// ---------- i18n --------------------------
						XsltArgumentList xsltI18nArgs = new XsltArgumentList();
						xsltI18nArgs.AddParam("lang", "", this.Language);
						XslTransform xslt = new XslTransform();
						//Load the i18n stylesheet.
						xslt.Load(_i18nXsl);
						XPathDocument xmlDoc;
						xmlDoc = new XPathDocument(xmlReader);
						XmlTextWriter writerFinal = new XmlTextWriter(Path.Combine(this.Todir, this.OutFilename), System.Text.Encoding.GetEncoding("utf-8"));
						// Apply the second transform to xmlReader to final ouput
						xslt.Transform(xmlDoc, xsltI18nArgs, writerFinal);

					}

				}
				else {
					StringReader stream;
					XmlTextReader reader = null;

					try {
#if ECHO_MODE
							Console.WriteLine ("Initializing StringReader ...");
#endif

						// create the index.html
						stream = new StringReader("<xsl:stylesheet xmlns:xsl='http://www.w3.org/1999/XSL/Transform' version='1.0' >" +
							"<xsl:output method='html' indent='yes' encoding='utf-8'/>" +
							"<xsl:include href=\"" + Path.Combine(_nantLocation, "NUnit-Frame.xsl") + "\"/>" +
							"<xsl:template match=\"test-results\">" +
							"   <xsl:call-template name=\"index.html\"/>" +
							" </xsl:template>" +
							" </xsl:stylesheet>");
						this.Write(stream, Path.Combine(this.Todir, this.OutFilename));

						// create the stylesheet.css
						stream = new StringReader("<xsl:stylesheet xmlns:xsl='http://www.w3.org/1999/XSL/Transform' version='1.0' >" +
							"<xsl:output method='html' indent='yes' encoding='utf-8'/>" +
							"<xsl:include href=\"" + Path.Combine(_nantLocation, "NUnit-Frame.xsl") + "\"/>" +
							"<xsl:template match=\"test-results\">" +
							"   <xsl:call-template name=\"stylesheet.css\"/>" +
							" </xsl:template>" +
							" </xsl:stylesheet>");
						this.Write(stream, Path.Combine(this.Todir, "stylesheet.css"));

						// create the overview-summary.html at the root 
						stream = new StringReader("<xsl:stylesheet xmlns:xsl='http://www.w3.org/1999/XSL/Transform' version='1.0' >" +
							"<xsl:output method='html' indent='yes' encoding='utf-8'/>" +
							"<xsl:include href=\"" + Path.Combine(_nantLocation, "NUnit-Frame.xsl") + "\"/>" +
							"<xsl:template match=\"test-results\">" +
							"    <xsl:call-template name=\"overview.packages\"/>" +
							" </xsl:template>" +
							" </xsl:stylesheet>");
						this.Write(stream, Path.Combine(this.Todir, "overview-summary.html"));


						// create the allclasses-frame.html at the root 
						stream = new StringReader("<xsl:stylesheet xmlns:xsl='http://www.w3.org/1999/XSL/Transform' version='1.0' >" +
							"<xsl:output method='html' indent='yes' encoding='utf-8'/>" +
							"<xsl:include href=\"" + Path.Combine(_nantLocation, "NUnit-Frame.xsl") + "\"/>" +
							"<xsl:template match=\"test-results\">" +
							"    <xsl:call-template name=\"all.classes\"/>" +
							" </xsl:template>" +
							" </xsl:stylesheet>");
						this.Write(stream, Path.Combine(this.Todir, "allclasses-frame.html"));

						// create the overview-frame.html at the root
						stream = new StringReader("<xsl:stylesheet xmlns:xsl='http://www.w3.org/1999/XSL/Transform' version='1.0' >" +
							"<xsl:output method='html' indent='yes' encoding='utf-8'/>" +
							"<xsl:include href=\"" + Path.Combine(_nantLocation, "NUnit-Frame.xsl") + "\"/>" +
							"<xsl:template match=\"test-results\">" +
							"    <xsl:call-template name=\"all.packages\"/>" +
							" </xsl:template>" +
							" </xsl:stylesheet>");
						this.Write(stream, Path.Combine(this.Todir, "overview-frame.html"));

						// Create directory
						string path = "";

						//--- Change 11/02/2003 -- remove
						//XmlDocument doc = new XmlDocument();
						//doc.Load("result.xml"); _FileSetSummary
						//---
						XPathNavigator xpathNavigator = _FileSetSummary.CreateNavigator(); //doc.CreateNavigator();

						// Get All the test suite containing test-case.
						XPathExpression expr = xpathNavigator.Compile("//test-suite[(child::results/test-case)]");

						XPathNodeIterator iterator = xpathNavigator.Select(expr);
						string directory = "";
						string testSuiteName = "";

						while( iterator.MoveNext() ) {
							XPathNavigator xpathNavigator2 = iterator.Current;
							testSuiteName = iterator.Current.GetAttribute("name", "");

#if ECHO_MODE
									Console.WriteLine("Test case : "+testSuiteName);   
#endif


							// Get get the path for the current test-suite.
							XPathNodeIterator iterator2 = xpathNavigator2.SelectAncestors("", "", true);
							path = "";
							string parent = "";
							int parentIndex = -1;

							while( iterator2.MoveNext() ) {
								directory = iterator2.Current.GetAttribute("name", "");
								if( directory != "" && directory.IndexOf(".dll") < 0 ) {
									path = directory + "/" + path;
								}
								if( parentIndex == 1 )
									parent = directory;
								parentIndex++;
							}
							Directory.CreateDirectory(Path.Combine(this.Todir, path));// path = xx/yy/zz

#if ECHO_MODE
									Console.WriteLine("path="+path+"\n");   
#endif

#if ECHO_MODE
									Console.WriteLine("parent="+parent+"\n");   
#endif

							// Build the "testSuiteName".html file
							// Correct MockError duplicate testName !
							// test-suite[@name='MockTestFixture' and ancestor::test-suite[@name='Assemblies'][position()=last()]] 

							stream = new StringReader("<xsl:stylesheet xmlns:xsl='http://www.w3.org/1999/XSL/Transform' version='1.0' >" +
								"<xsl:output method='html' indent='yes' encoding='utf-8'/>" +
								"<xsl:include href=\"" + Path.Combine(_nantLocation, "NUnit-Frame.xsl") + "\"/>" +
								"<xsl:template match=\"/\">" +
								"	<xsl:for-each select=\"//test-suite[@name='" + testSuiteName + "' and ancestor::test-suite[@name='" + parent + "'][position()=last()]]\">" +
								"		<xsl:call-template name=\"test-case\">" +
								"			<xsl:with-param name=\"dir.test\">" + String.Join(".", path.Split('/')) + "</xsl:with-param>" +
								"		</xsl:call-template>" +
								"	</xsl:for-each>" +
								" </xsl:template>" +
								" </xsl:stylesheet>");
							this.Write(stream, Path.Combine(Path.Combine(this.Todir, path), testSuiteName + ".html"));

#if ECHO_MODE
									Console.WriteLine("dir="+this.Todir+path+" Generate "+testSuiteName+".html\n");   
#endif
						}

#if ECHO_MODE
								Console.WriteLine ("Processing ...");
								Console.WriteLine ();	
#endif

					}

					catch( Exception e ) {
						Console.WriteLine("Exception: {0}", e.ToString());
					}

					finally {
#if ECHO_MODE
							Console.WriteLine();
							Console.WriteLine("Processing of stream complete.");	
#endif

						// Finished with XmlTextReader
						if( reader != null )
							reader.Close();
					}
				}
			}
			catch( Exception e ) {
				throw new Exception(e.Message);
			}
		}


		/// <summary>
		/// Initializes the XmlDocument instance
		/// used to summarize the test results
		/// </summary>
		/// <returns></returns>
		private XmlDocument CreateSummaryXmlDoc()
		{
			XmlDocument doc = new XmlDocument();
			XmlElement root = doc.CreateElement("testsummary");
			root.SetAttribute("created", DateTime.Now.ToString());
			doc.AppendChild(root);

			return doc;
		}

		/// <summary>
		/// Builds an XsltArgumentList with all
		/// the properties defined in the 
		/// current project as XSLT parameters.
		/// </summary>
		/// <returns></returns>
		private XsltArgumentList GetPropertyList()
		{
			XsltArgumentList args = new XsltArgumentList();

#if ECHO_MODE
		Console.WriteLine();
		Console.WriteLine("XsltArgumentList");	
#endif

			//            foreach( DictionaryEntry entry in Project.Properties ) {
			//#if ECHO_MODE
			//                Console.WriteLine();
			//                Console.WriteLine("Project.Properties :"+(string)entry.Key+"="+(string)entry.Value);	
			//#endif

			//                if( (string)entry.Value != null ) {
			//                    //Patch from Christoph Walcher 
			//                    try {
			//                        args.AddParam((string)entry.Key, "", (string)entry.Value);
			//                    }
			//                    catch( ArgumentException aex ) {
			//                        Console.WriteLine("Invalid Xslt parameter {0}", aex);
			//                    }
			//                }
			//            }

			// Add argument to the C# XML comment file
			args.AddParam("summary.xml", "", _tempXmlFileSummarie);
			// Add open.description argument 
			args.AddParam("open.description", "", this.OpenDescription);

			return args;
		}


		private void Write(StringReader stream, string fileName)
		{
			XmlTextReader reader = null;

			// Load the XmlTextReader from the stream
			reader = new XmlTextReader(stream);
			XslTransform xslTransform = new XslTransform();
			//Load the stylesheet from the stream.
			xslTransform.Load(reader);

			XPathDocument xmlDoc;
			//xmlDoc = new XPathDocument("result.xml");

			// xmlReader hold the first transformation 
			XmlReader xmlReader = xslTransform.Transform(_FileSetSummary, _xsltArgs);//(xmlDoc, _xsltArgs);

			// ---------- i18n --------------------------
			XsltArgumentList xsltI18nArgs = new XsltArgumentList();
			xsltI18nArgs.AddParam("lang", "", this.Language);


			XslTransform xslt = new XslTransform();

			//Load the stylesheet.
			xslt.Load(_i18nXsl);

			xmlDoc = new XPathDocument(xmlReader);

			XmlTextWriter writerFinal = new XmlTextWriter(fileName, System.Text.Encoding.GetEncoding("utf-8"));
			// Apply the second transform to xmlReader to final ouput
			xslt.Transform(xmlDoc, xsltI18nArgs, writerFinal);

			xmlReader.Close();
			writerFinal.Close();

		}

	}
	#endregion


	
}
