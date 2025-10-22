using System.Data;
using System.Text;
using UBS.Commons.Report.Common;
using UBS.Commons.Report.Common.Constants;
using UBS.Commons.Report.DTO;
using UBS.Commons.Report.Logics;
using UBS.Commons.Report.Logics.CreateReport;
using UBS.Core.Attributes;
using UBS.Core.BusinessLogic;
using UBS.Core.Extensions;
using UBS.Core.Utilities;
using UBS.Core.Utilities.DateCommons;
using UBS.KW21.Commons.Entities.KW21;
using UBS.KW21.Commons.Logics;
using UBS.KW21.Keiri.Logics.Commons;
using UBS.KW21.Keiri.Report.API.Constants.K0010;
using UBS.KW21.Keiri.Report.API.Constants.K001003;

namespace UBS.KW21.Keiri.Report.API.Logics
{
    /// <summary>
    /// 予算伝票モニタリスト出力（K001003）の帳票作成ロジック
    /// </summary>
    /// <remarks>
    /// <seealso cref="ICreateReportLogic"/>に定義されている共通の帳票データ作成処理メソッドを帳票ごとのロジックで実装する。
    /// </remarks>
    [Export(typeof(IK001003Logic))]
    public class K001003Logic : UBSLogic, ICreateReportLogic, IK001003Logic
    {
        #region メンバー変数定義

        /// <summary>行番号</summary>
        private int _lineNumber;

        /// <summary>備考グループキー</summary>
        private int _bikoGroupKey;

        /// <summary>出力日</summary>
        private string _outputDate = string.Empty;

        /// <summary>出力時間</summary>
        private string _outputTime = string.Empty;

        /// <summary>帳票共通ロジック</summary>
        [Import]
        public IReportFileLogic ReportFileLogic = null!;

        /// <summary><see cref="ShibuLogic"/></summary>
        [Import]
        public IShibuLogic ShibuLogic = null!;

        /// <summary><see cref="KetsugishoLogic"/></summary>
        [Import]
        public IKetsugishoLogic KetsugishoLogic = null!;

        #endregion

        #region コンストラクター定義

        /// <summary>
        /// コンストラクタです。
        /// </summary>
        public K001003Logic()
        {
        }

        #endregion

        #region プロパティ定義

        /// <summary>Kk19bikoTable</summary>
        [Import]
        public KW21Repository<KK19BIKO> Kk19bikoTable { get; set; } = null!;

        /// <summary>VKYosanMonitorTable</summary>
        [Import]
        public KW21Repository<VKYosanMonitor> VKYosanMonitorTable { get; set; } = null!;

        #endregion

        #region static メソッド定義

        /// <summary>
        /// CSVデータを取得します。
        /// </summary>
        /// <para><paramref name="printDataList"/> を、CSVデータ <see cref="YosanCsv"/> に加工します。</para>
        /// <para><see cref="List{YosanCsv}"/> に格納し、戻り値に設定します。</para>
        /// <param name="printDataList">予算伝票モニタリスト出力の帳票データリスト</param>
        /// <param name="token">キャンセルトークン</param>
        /// <returns><see cref="List{YosanCsv}"/></returns>
        static private List<YosanCsv> GetCsvData(List<PrintData> printDataList, CancellationToken token)
        {
            var csvDataList = new List<YosanCsv>();
            foreach (var r in printDataList)
            {
                token.ThrowIfCancellationRequested();
                var csvData = new YosanCsv
                {
                    LineNumber = r.LineNumber,
                    ShibuCode = r.HeaderShibuCode,
                    ShibuCodeFrom = r.VKYosanMonitor.frSIBUCODE,
                    ShushiKubunFrom = r.VKYosanMonitor.frSHUSIKBN,
                    KanFrom = r.VKYosanMonitor.frKANLVL,
                    KoFrom = r.VKYosanMonitor.frKOULVL,
                    MokuFrom = r.VKYosanMonitor.frMOKLVL,
                    DenpyoNumber = r.DenpyoNumber + r.DenpyoNumberBranch,
                    HatsugiDate = r.VKYosanMonitor.HATUGIYMD,
                    Nengo = !string.IsNullOrWhiteSpace(r.ShoriNendo) ? K0010StringLogic.SubString(r.ShoriNendo, 0, 2) : string.Empty.PadLeft(4),
                    Nendo = !string.IsNullOrWhiteSpace(r.ShoriNendo) ? K0010StringLogic.SubString(r.ShoriNendo, 2).PadLeft(2) : string.Empty.PadLeft(2),
                    ShibuNameFrom = K0010StringLogic.ConvertStringNullToEmpty(r.VKYosanMonitor.frSIBUNAME),
                    MokuNameFrom = K0010StringLogic.ConvertStringNullToEmpty(r.VKYosanMonitor.frMOKNAM),
                    Tekiyo = K0010StringLogic.ConvertStringNullToEmpty(r.VKYosanMonitor.TEKIYOU),
                    KetsugiDate = r.VKYosanMonitor.KESAIYMD,
                    TantoshaId = K0010StringLogic.ConvertStringNullToEmpty(r.VKYosanMonitor.USERCODE),
                    SuitoYoteiDate = r.VKYosanMonitor.SUITOYTD,
                    UkagaiKubunCodeUkagaiKubunSubCode = r.UkagaiKubunUkagaiKubunSub,
                    UkagaiKubunName = K0010StringLogic.ConvertStringNullToEmpty(r.VKYosanMonitor.UKAGAINAM),
                    ShibuCodeTo = r.VKYosanMonitor.toSIBUCODE,
                    ShushiKubunTo = r.VKYosanMonitor.toSHUSIKBN,
                    KanTo = r.VKYosanMonitor.toKANLVL,
                    KoTo = r.VKYosanMonitor.toKOULVL,
                    MokuTo = r.VKYosanMonitor.toMOKLVL,
                    Kingaku = r.VKYosanMonitor.YOSANGAKU,
                    SuitoDate = r.VKYosanMonitor.SUITOYMD,
                    ShibuNameTo = K0010StringLogic.ConvertStringNullToEmpty(r.VKYosanMonitor.toSIBUNAME),
                    MokuNameTo = K0010StringLogic.ConvertStringNullToEmpty(r.VKYosanMonitor.toMOKNAM),
                    BikoJodan = r.Biko1,
                    BikoGedan = r.Biko2
                };

                csvDataList.Add(csvData);
            }

            return csvDataList;
        }

        /// <summary>
        /// 科目コード設定します。
        /// </summary>
        /// <param name="shushiKubun">収支区分</param>
        /// <param name="kan">款コード</param>
        /// <param name="ko">項コード</param>
        /// <param name="moku">目コード</param>
        /// <returns>科目コード</returns>
        private static string SetKamokuCode(string? shushiKubun, int? kan, int? ko, int? moku)
        {
            if (string.IsNullOrWhiteSpace(shushiKubun))
            {
                return string.Empty;
            }

            var joionString = new string[]
            {
                K0010StringLogic.ConvertStringNullToEmpty(shushiKubun),
                kan.HasValue ? StringCommon.ToString((int)kan) : "0",
                ko.HasValue ? StringCommon.ToString((int)ko) : "0",
                moku.HasValue ? StringCommon.ToString((int)moku) : "0"
            };

            return string.Join(OutputWord.Hyphen, joionString);
        }

        #endregion

        #region メソッド定義

        /// <summary>
        /// 帳票データを作成します。
        /// </summary>
        /// <remarks>
        /// <para>帳票作成のメイン処理です。</para>
        /// <para><see cref="GetPrintData(ReportCreateDataDTO, string, List{string}, CancellationToken)"/> を使用し、帳票データを取得します。</para>
        /// <para><see cref="GetXmlData(List{PrintData}, CancellationToken)"/> を使用し、XMLデータを取得します。</para>
        /// <para><see cref="GetCsvData(List{PrintData}, CancellationToken)"/> を使用し、CSVデータを取得します。</para>
        /// <para><see cref="WritePrintDataFiles(string, string, string, ActiveReportPrintData, List{YosanCsv})"/> を使用し、帳票ファイルを作成・登録します。</para>
        /// </remarks>
        /// <param name="printId">印刷実行管理ID</param>
        /// <param name="reportCreateDataDTO">印刷要求DTO</param>
        /// <param name="outputFolder">出力フォルダ</param>
        /// <param name="token">キャンセルトークン</param>
        public void CreateReport(string printId, ReportCreateDataDTO reportCreateDataDTO, string outputFolder, CancellationToken token)
        {
            // 出力日時設定
            this._outputDate = DateCommon.SeirekiToWarekiDate(DateTime.Now.ToString(DateFormat.Yyyymmdd), DateFormat.EraDateFormatsKey.GeeNenMmGatsuDdNichi);
            this._outputTime = DateTime.Now.ToString(DateFormat.HColonMmColonSs);

            // 印刷パラメータの支部、伺い区分リストを取得
            string? sibuCodePara = reportCreateDataDTO.PrintParameter[Condition.ShibuCode];
            string? ukagaiKubunPara = reportCreateDataDTO.PrintParameter[Condition.UkagaiKubun];
            List<string> shibuCodeList = new();
            List<string> ukagaiKubunList = new();
            if (string.IsNullOrEmpty(sibuCodePara))
            {
                // 支部情報の取得
                var shibuJoho = this.ShibuLogic.GetUnyoShibuRecord();
                shibuCodeList.AddRange(shibuJoho.Select(e => e.SIBUCODE));
            }
            else
            {
                shibuCodeList.AddRange(sibuCodePara.Split(K001003Common.Comma).ToList());
            }

            if (!string.IsNullOrEmpty(ukagaiKubunPara))
            {
                ukagaiKubunList.AddRange(ukagaiKubunPara.Split(K001003Common.Comma).ToList());
            }

            var printDataList = new List<PrintData>();
            this._bikoGroupKey = 0;
            foreach (var shibuCode in shibuCodeList.OrderBy(e => e))
            {
                // 帳票データの取得
                this._lineNumber = 0;
                printDataList.AddRange(this.GetPrintData(reportCreateDataDTO, shibuCode, ukagaiKubunList, token));
            }

            // XMLデータの取得
            var xmlData = this.GetXmlData(printDataList, token);

            // CSVデータの取得
            var csvData = K001003Logic.GetCsvData(printDataList, token);

            // ファイル出力
            this.WritePrintDataFiles(printId, reportCreateDataDTO.ReportId, outputFolder, xmlData, csvData);
        }

        /// <summary>
        /// 帳票データを取得します。
        /// </summary>
        /// <remarks>
        /// <para>印刷パラメータの条件で、<see cref="VKYosanMonitorTable"/> を抽出します。</para>
        /// <para><see cref="SetPrintData"/> を使用して抽出データを帳票データに加工します。</para>
        /// <para><see cref="List{PrintData}"/>に格納し、戻り値に設定します。</para>
        /// </remarks>
        /// <param name="condition">抽出条件</param>
        /// <param name="shibuCode">支部コード</param>
        /// <param name="ukagaiKubunList">伺い区分リスト</param>
        /// <param name="token">キャンセルトークン</param>
        /// <returns><see cref="List{PrintData}"/></returns>
        private List<PrintData> GetPrintData(ReportCreateDataDTO condition, string shibuCode, List<string> ukagaiKubunList, CancellationToken token)
        {
            // 印刷パラメータ設定
            string userCode = condition.PrintParameter[Condition.UserCode];
            string kessaiKubun = condition.PrintParameter[Condition.KessaiKubun];
            string shoriNendo = condition.PrintParameter[Condition.ShoriNendo];
            string hatsugiDateFrom = condition.PrintParameter[Condition.HatsugiDateFrom];
            string hatsugiDateTo = condition.PrintParameter[Condition.HatsugiDateTo];
            string suitoYoteiDateFrom = condition.PrintParameter[Condition.SuitoYoteiDateFrom];
            string suitoYoteiDateTo = condition.PrintParameter[Condition.SuitoYoteiDateTo];
            string suitoDateFrom = condition.PrintParameter[Condition.SuitoDateFrom];
            string suitoDateTo = condition.PrintParameter[Condition.SuitoDateTo];
            string dempyoNumberFrom = condition.PrintParameter[Condition.DempyoNumberFrom];
            string dempyoNumberTo = condition.PrintParameter[Condition.DempyoNumberTo];
            string kanjoKubunCode = condition.PrintParameter[Condition.KanjoKubunCode];

            // 勘定区分リスト：一般
            var ippanList = new List<string> { K0010KeiriCommon.ShushiKubun.IppanKanjoShunyuKamoku, K0010KeiriCommon.ShushiKubun.IppanKanjoShishutsuKamoku };

            // 勘定区分リスト：介護
            var kaigoList = new List<string> { K0010KeiriCommon.ShushiKubun.KaigoKanjoShunyuKamoku, K0010KeiriCommon.ShushiKubun.KaigoKanjoShishutsuKamoku };

            // 勘定区分リスト：子ども
            var childList = new List<string> { K0010KeiriCommon.ShushiKubun.ChildKanjoShunyuKamoku, K0010KeiriCommon.ShushiKubun.ChildKanjoShishutsuKamoku };

            // 伺い区分リスト：前年度
            var ukagaiZennenList = ukagaiKubunList.Where(e => e.Substring(1, 3) == K0010KeiriCommon.UkagaiKubun.ZennendoShushiZankinIchijiJutoKetsugisho);

            // 伺い区分リスト：その他
            var ukagaiSonotaList = ukagaiKubunList.Where(e => e.Substring(1, 3) != K0010KeiriCommon.UkagaiKubun.ZennendoShushiZankinIchijiJutoKetsugisho);

            // 予算伝票モニタリストビューを取得
            var data = this.VKYosanMonitorTable
                .Where(e => e.frSIBUCODE! == shibuCode || e.toSIBUCODE! == shibuCode)
                .WhereIf(!string.IsNullOrWhiteSpace(hatsugiDateFrom), e => string.Compare(e.HATUGIYMD, hatsugiDateFrom) >= 0)
                .WhereIf(!string.IsNullOrWhiteSpace(hatsugiDateTo), e => string.Compare(e.HATUGIYMD, hatsugiDateTo) <= 0)
                .WhereIf(!string.IsNullOrWhiteSpace(suitoYoteiDateFrom), e => string.Compare(e.SUITOYTD, suitoYoteiDateFrom) >= 0)
                .WhereIf(!string.IsNullOrWhiteSpace(suitoYoteiDateTo), e => string.Compare(e.SUITOYTD, suitoYoteiDateTo) <= 0)
                .WhereIf(!string.IsNullOrWhiteSpace(suitoDateFrom), e => string.Compare(e.SUITOYMD, suitoDateFrom) >= 0)
                .WhereIf(!string.IsNullOrWhiteSpace(suitoDateTo), e => string.Compare(e.SUITOYMD, suitoDateTo) <= 0)
                .WhereIf(!string.IsNullOrWhiteSpace(dempyoNumberFrom), e => e.DENPYOBG >= StringCommon.ToInt(dempyoNumberFrom))
                .WhereIf(!string.IsNullOrWhiteSpace(dempyoNumberTo), e => e.DENPYOBG <= StringCommon.ToInt(dempyoNumberTo))
                .WhereIf(!string.IsNullOrWhiteSpace(shoriNendo), e => e.SHORINEN == shoriNendo)
                .WhereIf(!string.IsNullOrWhiteSpace(userCode), e => e.USERCODE == userCode)
                .WhereIf(ukagaiZennenList.Any(), e => ukagaiZennenList.Contains(e.UKAGAIKB))
                .WhereIf(ukagaiSonotaList.Any(), e => ukagaiSonotaList.Contains(e.UKAGAIKB + e.UKAGAISUB))
                .WhereIf(kanjoKubunCode == K0010KeiriCommon.KanjoKubun.KenkoHokenKanjo, e => ippanList.Contains(e.frSHUSIKBN!) || ippanList.Contains(e.toSHUSIKBN!))
                .WhereIf(kanjoKubunCode == K0010KeiriCommon.KanjoKubun.KaigoHokenKanjo, e => kaigoList.Contains(e.frSHUSIKBN!) || kaigoList.Contains(e.toSHUSIKBN!))
                .WhereIf(kanjoKubunCode == K0010KeiriCommon.KanjoKubun.ChildHokenKanjo, e => childList.Contains(e.frSHUSIKBN!) || childList.Contains(e.toSHUSIKBN!))
                .WhereIf(kessaiKubun == K0010KeiriCommon.KessaiKubun.MiketsuBun, e => e.KESAIYMD == K001003Common.DateTimeNoSet)
                .WhereIf(kessaiKubun == K0010KeiriCommon.KessaiKubun.KessaiBun, e => e.KESAIYMD != K001003Common.DateTimeNoSet)
                .OrderBy(e => e.SHORINEN)
                .ThenBy(e => e.DENPYOBG)
                .ThenBy(e => e.DENPYOSUB)
                .ThenBy(e => e.HATUGIYMD)
                .ThenBy(e => e.UKAGAIKB)
                .ThenBy(e => e.UKAGAISUB)
                .ThenBy(e => e.frSIBUCODE)
                .ThenBy(e => e.toSIBUCODE)
                .ThenBy(e => e.SUITOYMD);

            // 支部名称の取得
            var shibuNameJoho = this.ShibuLogic.GetShibuRecordSearchName(shibuCode, string.Empty);
            var shibuName = K0010StringLogic.ConvertStringNullToEmpty(shibuNameJoho.First().SKANJNAM);

            var printDataLiset = new List<PrintData>();
            foreach (var r in data)
            {
                token.ThrowIfCancellationRequested();

                var printData = this.SetPrintData(r);
                printData.HeaderShibuCode = shibuCode;
                printData.ShibuInfo = $"{shibuCode}{OutputWord.Colon}{shibuName}";

                printDataLiset.Add(printData);
            }

            return printDataLiset;
        }

        /// <summary>
        /// 帳票データを設定します。
        /// </summary>
        /// <remarks>
        /// 抽出レコード <paramref name="record"/> を、帳票データ <see cref="PrintData"/> に加工します。
        /// </remarks>
        /// <param name="record">抽出レコード</param>
        /// <returns><see cref="PrintData"/></returns>
        private PrintData SetPrintData(VKYosanMonitor record)
        {
            // 経理伝票備考を取得
            var bikoData = this.Kk19bikoTable
                .Where(e => e.SHORINEN == record.SHORINEN)
                .Where(e => e.DENPYOBG == record.DENPYOBG)
                .OrderBy(e => e.BIKOKBN);

            this._lineNumber++;
            this._bikoGroupKey++;
            var printData = new PrintData
            {
                VKYosanMonitor = record,
                LineNumber = this._lineNumber,
                BikoGroupKey = this._bikoGroupKey,
                DenpyoNumber = StringCommon.ToString(record.DENPYOBG).PadLeft(9),
                DenpyoNumberBranch = OutputWord.Hyphen + StringCommon.ToString(record.DENPYOSUB).PadLeft(5),
                UkagaiKubunUkagaiKubunSub = $"{record.UKAGAIKB}{OutputWord.Hyphen}{record.UKAGAISUB}",
                ShoriNendo = !string.IsNullOrWhiteSpace(record.SHORINEN) && !string.IsNullOrWhiteSpace(record.HATUGIYMD) ?
                    this.KetsugishoLogic.SeirekiToWarekiNendoByHatsugiDate(record.SHORINEN, record.HATUGIYMD, DateFormat.FiscalEraFormatsKey.Ge) : string.Empty,
                Biko1 = K0010StringLogic.ConvertStringNullToEmpty(bikoData.Where(e => e.BIKOSEQ == 1).FirstOrDefault()?.BIKO),
                Biko2 = K0010StringLogic.ConvertStringNullToEmpty(bikoData.Where(e => e.BIKOSEQ == 2).FirstOrDefault()?.BIKO)
            };

            return printData;
        }

        /// <summary>
        /// 帳票データを出力します。
        /// </summary>
        /// <remarks>
        /// <para><paramref name="xmlData"/> をXMLデータに出力します。</para>
        /// <para><paramref name="csvData"/> をCSVに出力します。</para>
        /// <para><see cref="ReportFileLogic.AddReportFile(string, string, string, string)"/> を使用し、作成したファイルを登録します。</para>
        /// </remarks>
        /// <param name="printId">帳票作成管理ID</param>
        /// <param name="programId">プログラムID</param>
        /// <param name="seq">SEQ</param>
        /// <param name="outputFolder">出力フォルダ</param>
        /// <param name="xmlData">帳票XMLデータ</param>
        /// <param name="csvData">CSVデータ</param>
        private void WritePrintDataFiles(
            string printId,
            string programId,
            string outputFolder,
            ActiveReportPrintData xmlData,
            List<YosanCsv> csvData)
        {
            // 出力ファイル名を取得
            var reportFileName = ReportFile.GetFileName(ReportFileTypeConstant.Report, printId, programId);
            var csvFileName = ReportFile.GetFileName(ReportFileTypeConstant.Csv, printId, programId);

            // 帳票データファイル(XML)を出力
            using (var sw = new StreamWriter(
                Path.Combine(outputFolder, reportFileName),
                false,
                Encoding.UTF8))
            {
                xmlData.Serialize(sw);
            }

            // 帳票データファイルを登録
            this.ReportFileLogic.AddReportFile(printId, outputFolder, reportFileName, ReportFileTypeConstant.Report);

            // CSVデータファイルを出力
            if (csvData.Any())
            {
                // CSVデータファイルを出力
                using (var sw =
                    new StreamWriter(Path.Combine(outputFolder, csvFileName),
                        false,
                        Encoding.GetEncoding(K001003Common.EncodingShiftJis)))
                {
                    var csvCommon = new CsvCommon();
                    csvCommon.WriteCSV<YosanCsv>(sw, csvData, true, true);
                }

                // CSVデータファイルを登録
                this.ReportFileLogic.AddReportFile(printId, outputFolder, csvFileName, ReportFileTypeConstant.Csv);
            }
        }

        /// <summary>
        /// XMLデータを取得します。
        /// </summary>
        /// <remarks>
        /// <paramref name="printDataList"/> を、XMLデータ <see cref="ActiveReportDataFields"/> に加工します。
        /// </remarks>
        /// <param name="printDataList">帳票データ</param>
        /// <param name="token">キャンセルトークン</param>
        /// <returns><see cref="ActiveReportPrintData"/></returns>
        private ActiveReportPrintData GetXmlData(List<PrintData> printDataList, CancellationToken token)
        {
            var ar = new ActiveReportPrintData(new string[]
            {
                ActiveReportDataFields.PageNumber,
                ActiveReportDataFields.KeyBiko,
                ActiveReportDataFields.Sibuinfo,
                ActiveReportDataFields.Outputymd,
                ActiveReportDataFields.Outputtime,
                ActiveReportDataFields.Denpyobg,
                ActiveReportDataFields.Denpyosub,
                ActiveReportDataFields.Hatugiymd,
                ActiveReportDataFields.Suitoytd,
                ActiveReportDataFields.Shorinen,
                ActiveReportDataFields.Ukagaikb,
                ActiveReportDataFields.Ukagainam,
                ActiveReportDataFields.Frsibucode,
                ActiveReportDataFields.Frsibuname,
                ActiveReportDataFields.Tosibucod,
                ActiveReportDataFields.Tosibuname,
                ActiveReportDataFields.Frkamokucd,
                ActiveReportDataFields.Frmoknam,
                ActiveReportDataFields.Tokamokucd,
                ActiveReportDataFields.Tomoknam,
                ActiveReportDataFields.Tekiyou,
                ActiveReportDataFields.Yosangaku,
                ActiveReportDataFields.Kesaiymd,
                ActiveReportDataFields.Suitoymd,
                ActiveReportDataFields.Usercode,
                ActiveReportDataFields.Biko1,
                ActiveReportDataFields.Biko2
            });

            int printHeight = 0;
            int pageNumber = 1;
            var prevSibuCode = string.Empty;
            foreach (var r in printDataList)
            {
                token.ThrowIfCancellationRequested();

                // １ページに表示する明細行の制御を行う。
                if (printHeight > K001003Common.MaxHeight || (!string.IsNullOrWhiteSpace(prevSibuCode) && prevSibuCode != r.HeaderShibuCode))
                {
                    // 以下の条件のいずれかに合致する場合、改ページを行う。
                    // ①：１ページに表示する高さの合計が既定値を超える場合
                    // ②：支部コードが変わった場合
                    printHeight = 0;
                    pageNumber++;
                    ar.NewPage();
                }

                // 支部コードが替わったら、ページ番号を1から振りなおす。
                if (!string.IsNullOrWhiteSpace(prevSibuCode) && prevSibuCode != r.HeaderShibuCode)
                {
                    pageNumber = 1;
                }

                ar.SetValue(ActiveReportDataFields.PageNumber, StringCommon.ToString(pageNumber));
                ar.SetValue(ActiveReportDataFields.KeyBiko, StringCommon.ToString(r.BikoGroupKey - 1));
                ar.SetValue(ActiveReportDataFields.Outputymd, this._outputDate);
                ar.SetValue(ActiveReportDataFields.Outputtime, this._outputTime);
                ar.SetValue(ActiveReportDataFields.Sibuinfo, K0010StringLogic.ConvertStringNullToEmpty(r.ShibuInfo));
                ar.SetValue(ActiveReportDataFields.Denpyobg, r.DenpyoNumber!);
                ar.SetValue(ActiveReportDataFields.Denpyosub, r.DenpyoNumberBranch!);
                ar.SetValue(ActiveReportDataFields.Hatugiymd, DateCommon.SeirekiToWarekiDate(r.VKYosanMonitor.HATUGIYMD, DateFormat.EraDateFormatsKey.AeeDotMmDotDd));
                ar.SetValue(ActiveReportDataFields.Suitoytd, DateCommon.SeirekiToWarekiDate(r.VKYosanMonitor.SUITOYTD, DateFormat.EraDateFormatsKey.AeeDotMmDotDd));
                ar.SetValue(ActiveReportDataFields.Shorinen, r.ShoriNendo + OutputWord.Nendo);
                ar.SetValue(ActiveReportDataFields.Ukagaikb, r.UkagaiKubunUkagaiKubunSub!);
                ar.SetValue(ActiveReportDataFields.Ukagainam, K0010StringLogic.ConvertStringNullToEmpty(r.VKYosanMonitor.UKAGAINAM));
                ar.SetValue(ActiveReportDataFields.Frsibucode, K0010StringLogic.ConvertAllZeoToEmpty(r.VKYosanMonitor.frSIBUCODE));
                ar.SetValue(ActiveReportDataFields.Frsibuname, K0010StringLogic.ConvertStringNullToEmpty(r.VKYosanMonitor.frSIBUNAME));
                ar.SetValue(ActiveReportDataFields.Tosibucod, K0010StringLogic.ConvertAllZeoToEmpty(r.VKYosanMonitor.toSIBUCODE));
                ar.SetValue(ActiveReportDataFields.Tosibuname, K0010StringLogic.ConvertStringNullToEmpty(r.VKYosanMonitor.toSIBUNAME));
                ar.SetValue(ActiveReportDataFields.Frkamokucd, K001003Logic.SetKamokuCode(r.VKYosanMonitor.frSHUSIKBN, r.VKYosanMonitor.frKANLVL, r.VKYosanMonitor.frKOULVL, r.VKYosanMonitor.frMOKLVL));
                ar.SetValue(ActiveReportDataFields.Frmoknam, K0010StringLogic.SubString(K0010StringLogic.ConvertStringNullToEmpty(r.VKYosanMonitor.frMOKNAM), 0, 20));
                ar.SetValue(ActiveReportDataFields.Tokamokucd, K001003Logic.SetKamokuCode(r.VKYosanMonitor.toSHUSIKBN, r.VKYosanMonitor.toKANLVL, r.VKYosanMonitor.toKOULVL, r.VKYosanMonitor.toMOKLVL));
                ar.SetValue(ActiveReportDataFields.Tomoknam, K0010StringLogic.SubString(K0010StringLogic.ConvertStringNullToEmpty(r.VKYosanMonitor.toMOKNAM), 0, 20));
                ar.SetValue(ActiveReportDataFields.Tekiyou, K0010StringLogic.ConvertStringNullToEmpty(r.VKYosanMonitor.TEKIYOU));
                ar.SetValue(ActiveReportDataFields.Yosangaku, r.VKYosanMonitor.YOSANGAKU.HasValue ? StringCommon.ToString((double)r.VKYosanMonitor.YOSANGAKU) : string.Empty);
                ar.SetValue(ActiveReportDataFields.Kesaiymd,
                    r.VKYosanMonitor.KESAIYMD != K001003Common.DateTimeNoSet ? DateCommon.SeirekiToWarekiDate(r.VKYosanMonitor.KESAIYMD, DateFormat.EraDateFormatsKey.AeeDotMmDotDd) : String.Empty);
                ar.SetValue(ActiveReportDataFields.Suitoymd,
                    r.VKYosanMonitor.SUITOYMD != K001003Common.DateTimeNoSet ? DateCommon.SeirekiToWarekiDate(r.VKYosanMonitor.SUITOYMD, DateFormat.EraDateFormatsKey.AeeDotMmDotDd) : String.Empty);
                ar.SetValue(ActiveReportDataFields.Usercode, K0010StringLogic.ConvertStringNullToEmpty(r.VKYosanMonitor.USERCODE));
                ar.SetValue(ActiveReportDataFields.Biko1, K0010StringLogic.ConvertStringNullToEmpty(r.Biko1));
                ar.SetValue(ActiveReportDataFields.Biko2, K0010StringLogic.ConvertStringNullToEmpty(r.Biko2));

                // 備考１・２のいずれかが設定されている場合、１ページに表示する高さの合計に「6」を加算する。
                if (!string.IsNullOrWhiteSpace(r.Biko1) || !string.IsNullOrWhiteSpace(r.Biko2))
                {
                    printHeight += K001003Common.FooterHeight;
                }

                prevSibuCode = r.HeaderShibuCode;
                printHeight += K001003Common.DetailHeight;
                ar.AddRow();
            }

            return ar;
        }

        #endregion

        #region ネストクラス(インナークラス)定義

        /// <summary>
        /// 予算伝票モニタリスト出力の帳票データ格納クラスです。
        /// </summary>
        internal class PrintData
        {
            /// <summary>行番号</summary>
            internal int LineNumber { get; set; }

            /// <summary>備考グループキー</summary>
            internal int BikoGroupKey { get; set; }

            /// <summary>ヘッダー支部コード</summary>
            internal string? HeaderShibuCode { get; set; }

            /// <summary>支部情報</summary>
            internal string? ShibuInfo { get; set; }

            /// <summary>伝票番号</summary>
            internal string? DenpyoNumber { get; set; }

            /// <summary>伝票番号枝番</summary>
            internal string? DenpyoNumberBranch { get; set; }

            /// <summary>伺い区分・伺い区分サブ</summary>
            internal string? UkagaiKubunUkagaiKubunSub { get; set; }

            /// <summary>処理年度（和暦）</summary>
            internal string? ShoriNendo { get; set; }

            /// <summary>備考１</summary>
            internal string? Biko1 { get; set; }

            /// <summary>備考２</summary>
            internal string? Biko2 { get; set; }

            /// <summary><see cref="VKYosanMonitor"/></summary>
            internal VKYosanMonitor VKYosanMonitor { get; set; } = null!;
        }

        /// <summary>
        /// 予算伝票モニタリスト出力のCsv格納クラスです。
        /// </summary>
        internal class YosanCsv
        {
            /// <summary>■行番号</summary>
            [CsvColumn(1, Name = "■行番号", NullValue = "")]
            public int? LineNumber { get; set; }

            /// <summary>支部コード</summary>
            [CsvColumn(2, Name = "支部コード", NullValue = "")]
            public string? ShibuCode { get; set; }

            /// <summary>支部コード･自</summary>
            [CsvColumn(3, Name = "支部コード･自", NullValue = "")]
            public string? ShibuCodeFrom { get; set; }

            /// <summary>収支区分(1:一般勘定収入、2:一般勘定支出、3:介護勘定収入、4:介護勘定支出)･自</summary>
            [CsvColumn(4, Name = "収支区分(1:一般勘定収入、2:一般勘定支出、3:介護勘定収入、4:介護勘定支出、5:子ども勘定収入、6:子ども勘定支出)･自", NullValue = "")]
            public string? ShushiKubunFrom { get; set; }

            /// <summary>款番号･自</summary>
            [CsvColumn(5, Name = "款番号･自", NullValue = "")]
            public int? KanFrom { get; set; }

            /// <summary>項番号･自</summary>
            [CsvColumn(6, Name = "項番号･自", NullValue = "")]
            public int? KoFrom { get; set; }

            /// <summary>目番号･自</summary>
            [CsvColumn(7, Name = "目番号･自", NullValue = "")]
            public int? MokuFrom { get; set; }

            /// <summary>伝票番号-伝票番号枝番</summary>
            [CsvColumn(8, Name = "伝票番号-伝票番号枝番", NullValue = "")]
            public string? DenpyoNumber { get; set; }

            /// <summary>発議日</summary>
            [CsvColumn(9, Name = "発議日", NullValue = "")]
            public string? HatsugiDate { get; set; }

            /// <summary>年号</summary>
            [CsvColumn(10, Name = "年号", NullValue = "")]
            public string? Nengo { get; set; }

            /// <summary>年度</summary>
            [CsvColumn(11, Name = "年度", NullValue = "")]
            public string? Nendo { get; set; }

            /// <summary>支部名称･自</summary>
            [CsvColumn(12, Name = "支部名称･自", NullValue = "")]
            public string? ShibuNameFrom { get; set; }

            /// <summary>目名称･自</summary>
            [CsvColumn(13, Name = "目名称･自", NullValue = "")]
            public string? MokuNameFrom { get; set; }

            /// <summary>摘要</summary>
            [CsvColumn(14, Name = "摘要", NullValue = "")]
            public string? Tekiyo { get; set; }

            /// <summary>決議日</summary>
            [CsvColumn(15, Name = "決議日", NullValue = "")]
            public string? KetsugiDate { get; set; }

            /// <summary>担当者ID</summary>
            [CsvColumn(16, Name = "担当者ID", NullValue = "")]
            public string? TantoshaId { get; set; }

            /// <summary>出納予定日</summary>
            [CsvColumn(17, Name = "出納予定日", NullValue = "")]
            public string? SuitoYoteiDate { get; set; }

            /// <summary>伺区分コード-伺区分サブコード</summary>
            [CsvColumn(18, Name = "伺区分コード-伺区分サブコード", NullValue = "")]
            public string? UkagaiKubunCodeUkagaiKubunSubCode { get; set; }

            /// <summary>伺区分名称</summary>
            [CsvColumn(19, Name = "伺区分名称", NullValue = "")]
            public string? UkagaiKubunName { get; set; }

            /// <summary>支部コード･至</summary>
            [CsvColumn(20, Name = "支部コード･至", NullValue = "")]
            public string? ShibuCodeTo { get; set; }

            /// <summary>収支区分(1:一般勘定収入、2:一般勘定支出、3:介護勘定収入、4:介護勘定支出)･至</summary>
            [CsvColumn(21, Name = "収支区分(1:一般勘定収入、2:一般勘定支出、3:介護勘定収入、4:介護勘定支出、5:子ども勘定収入、6:子ども勘定支出)･至", NullValue = "")]
            public string? ShushiKubunTo { get; set; }

            /// <summary>款番号･至</summary>
            [CsvColumn(22, Name = "款番号･至", NullValue = "")]
            public int? KanTo { get; set; }

            /// <summary>項番号･至</summary>
            [CsvColumn(23, Name = "項番号･至", NullValue = "")]
            public int? KoTo { get; set; }

            /// <summary>目番号･至</summary>
            [CsvColumn(24, Name = "目番号･至", NullValue = "")]
            public int? MokuTo { get; set; }

            /// <summary>金額</summary>
            [CsvColumn(25, Name = "金額", NullValue = "")]
            public double? Kingaku { get; set; }

            /// <summary>出納日</summary>
            [CsvColumn(26, Name = "出納日", NullValue = "")]
            public string? SuitoDate { get; set; }

            /// <summary>支部名称･至</summary>
            [CsvColumn(27, Name = "支部名称･至", NullValue = "")]
            public string? ShibuNameTo { get; set; }

            /// <summary>目名称･至</summary>
            [CsvColumn(28, Name = "目名称･至", NullValue = "")]
            public string? MokuNameTo { get; set; }

            /// <summary>備考上段</summary>
            [CsvColumn(29, Name = "備考上段", NullValue = "")]
            public string? BikoJodan { get; set; }

            /// <summary>備考下段</summary>
            [CsvColumn(30, Name = "備考下段", NullValue = "")]
            public string? BikoGedan { get; set; }
        }

        #endregion
    }
}
