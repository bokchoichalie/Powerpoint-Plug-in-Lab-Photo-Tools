using System;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace LabPhotoTools
{
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    [Guid("7E93D507-A159-4E38-9D53-A7594C858BD7")]
    [ComVisible(true)]
    [ProgId("LabPhotoTools.Connect")]
    public partial class Connect : IDTExtensibility2, IRibbonExtensibility
    {
        private object application;
        private Form currentForm;
        public void OnConnection(object app, int mode, object instance, ref Array custom) { application = app; }
        public void OnDisconnection(int mode, ref Array custom) { if (currentForm != null) currentForm.Close(); ribbonUI = null; application = null; }
        public void OnAddInsUpdate(ref Array custom) { }
        public void OnStartupComplete(ref Array custom) { }
        public void OnBeginShutdown(ref Array custom) { if (currentForm != null) currentForm.Close(); ribbonUI = null; }

        public string GetCustomUI(string ribbonId)
        {
            StringBuilder xml = new StringBuilder(@"<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui' onLoad='OnRibbonLoad'>
              <ribbon><tabs><tab id='labPhotoTab' label='Lab Photo Tools'>
                <group id='labPhotoImage' label='사진'>
                  <button id='labPhotoEdit' label='사진 회전' size='large' getImage='GetIcon' onAction='OpenImages' screentip='사진 회전' supertip='0.1° 단위로 회전하고 빈 모서리를 자릅니다. 원본 슬라이드의 복사본에 적용합니다.'/>
                  <button id='labPhotoBackground' label='배경지우기' size='large' getImage='GetIcon' onAction='OpenBackground' screentip='사진 배경지우기' supertip='선택한 사진의 배경을 이 PC에서 제거합니다. 미리보기를 확인하고 복사본에 적용하세요.'/>
                </group>
                <group id='labPhotoLayout' label='배치'>
                  <button id='labPhotoGrid' label='자동 배열' size='large' getImage='GetIcon' onAction='OpenLayout' supertip='선택한 항목 중 사진의 열 수, 너비와 간격을 직접 지정합니다. 도형과 텍스트는 그대로 둡니다.'/>
                  <button id='labPhotoSpacing' label='간격 조절' size='large' getImage='GetIcon' onAction='OpenSpacing' supertip='선택한 항목 중 사진의 가로 및 세로 간격을 조절합니다. 도형과 텍스트는 그대로 둡니다.'/>
                  <button id='labPhotoMagic' label='알아서 배열' size='large' getImage='GetIcon' onAction='ArrangeSmart' screentip='선택한 사진 또는 현재 슬라이드 사진을 배열' supertip='사진이 선택되어 있으면 그 사진만, 선택이 없으면 현재 슬라이드의 모든 사진을 비율에 맞춰 앞 행부터 채웁니다. 도형과 텍스트는 그대로 둡니다.'/>
                </group>
                <group id='labPhotoNumbers' label='빠른 번호 매기기'>");
            string[] names = { "정사각형", "원", "(알파벳)", "숫자)", "알파벳)" };
            string[] styles = { "square", "circle", "paren", "suffix", "alphaSuffix" };
            for (int s = 0; s < styles.Length; s++)
            {
                string style = styles[s];
                if (s > 0) xml.AppendFormat("<separator id='numberSeparator_{0}'/>", style);
                xml.AppendFormat("<box id='numberBox_{0}' boxStyle='vertical'>", style);
                // A three-row, four-column block keeps all eleven presets and
                // the next-number action directly accessible on the ribbon.
                for (int row = 0; row < 3; row++)
                {
                    xml.AppendFormat("<buttonGroup id='numberRow_{0}_{1}'>", style, row);
                    for (int column = 0; column < 4; column++)
                    {
                        int n = row * 4 + column;
                        if (n <= 10)
                            xml.AppendFormat("<button id='number_{0}_{1}' tag='{0}:{1}' label='{2}' showImage='false' onAction='AddNumber' screentip='{3} {4} 추가'/>", style, n, NumberLabels.Caption(style, n), names[s], NumberLabels.Text(style, n));
                        else
                            xml.AppendFormat("<button id='number_{0}_next' tag='{0}:next' label='{2}' showImage='false' onAction='AddNumber' screentip='{1} 다음 번호' supertip='{3}'/>", style, names[s], NumberLabels.IsAlphabet(style) ? "L+" : "11+", NumberLabels.IsAlphabet(style) ? "이 슬라이드에서 같은 형식의 마지막 알파벳 다음 값을 추가합니다. L부터 자동 증가합니다." : "이 슬라이드에서 같은 형식의 가장 큰 번호 다음 숫자를 추가합니다. 11부터 자동 증가합니다.");
                    }
                    xml.Append("</buttonGroup>");
                }
                xml.Append("</box>");
            }
            xml.Append(@"<separator id='numberSettingsSeparator'/>
                <box id='numberFormatBox' boxStyle='vertical'>
                  <dropDown id='labNumberFont' label='글꼴' sizeString='Times New Roman' showItemImage='false' getItemCount='GetNumberFontCount' getItemLabel='GetNumberFontLabel' getSelectedItemIndex='GetNumberFontIndex' onAction='SetNumberFont' screentip='새 번호 라벨의 글꼴' supertip='선택하면 즉시 저장되고 이후 만드는 번호에 적용됩니다.'/>
                  <editBox id='labNumberFontSize' label='크기 (pt)' sizeString='000.0' maxLength='8' getText='GetNumberFontSize' onChange='SetNumberFontSize' screentip='새 번호 라벨의 글자 크기' supertip='1~400 pt를 입력하고 Enter를 누르세요. 이후 만드는 번호에 적용됩니다.'/>
                  <box id='numberColorRow' boxStyle='horizontal'>
                    <gallery id='labNumberColor' label='글자 색' showLabel='true' getImage='GetNumberColorImage' columns='8' rows='4' itemWidth='24' itemHeight='24' showItemLabel='false' getItemCount='GetNumberColorCount' getItemLabel='GetNumberColorLabel' getItemImage='GetNumberColorItemImage' onAction='SetNumberColor' screentip='새 번호 라벨의 글자 색' supertip='색을 선택하면 이후 만드는 번호에 적용됩니다.'/>
                    <editBox id='labNumberColorHex' label='#' sizeString='FFFFFF' maxLength='7' getText='GetNumberColorHex' onChange='SetNumberColorHex' screentip='색상 코드 직접 입력' supertip='원하는 색의 여섯 자리 코드를 입력하고 Enter를 누르세요. 예: 검정 000000, 흰색 FFFFFF, 빨강 FF0000.'/>
                  </box>
                </box>
                <separator id='numberBoxSettingsSeparator'/>
                <box id='numberShapeFormatBox' boxStyle='vertical'>
                  <box id='numberBorderRow' boxStyle='horizontal'>
                    <gallery id='labNumberBorderColor' label='테두리 색' showLabel='true' getImage='GetNumberBoxColorImage' columns='8' rows='5' itemWidth='24' itemHeight='24' showItemLabel='false' getItemCount='GetNumberBoxColorCount' getItemLabel='GetNumberBoxColorLabel' getItemImage='GetNumberBoxColorItemImage' onAction='SetNumberBoxColor' supertip='이후 만드는 모든 라벨에 적용합니다. 기본값 또는 없음도 선택할 수 있습니다.'/>
                    <editBox id='labNumberBorderHex' label='#' sizeString='FFFFFF' maxLength='7' getText='GetNumberBoxColorHex' onChange='SetNumberBoxColorHex' supertip='여섯 자리 색상 코드, 기본 또는 없음을 입력하고 Enter를 누르세요.'/>
                  </box>
                  <editBox id='labNumberBorderWidth' label='테두리 (pt)' sizeString='00.00' maxLength='8' getText='GetNumberBorderWidth' onChange='SetNumberBorderWidth' supertip='0~20 pt. 0은 테두리 없음입니다. Enter로 확정하세요.'/>
                  <box id='numberFillRow' boxStyle='horizontal'>
                    <gallery id='labNumberFillColor' label='바탕색' showLabel='true' getImage='GetNumberBoxColorImage' columns='8' rows='5' itemWidth='24' itemHeight='24' showItemLabel='false' getItemCount='GetNumberBoxColorCount' getItemLabel='GetNumberBoxColorLabel' getItemImage='GetNumberBoxColorItemImage' onAction='SetNumberBoxColor' supertip='이후 만드는 모든 라벨에 적용합니다. 기본값 또는 없음(투명)도 선택할 수 있습니다.'/>
                    <editBox id='labNumberFillHex' label='#' sizeString='FFFFFF' maxLength='7' getText='GetNumberBoxColorHex' onChange='SetNumberBoxColorHex' supertip='여섯 자리 색상 코드, 기본 또는 없음을 입력하고 Enter를 누르세요.'/>
                  </box>
                </box></group>
                  <group id='labMeasurementGroup' label='치수측정'>
                    <button id='labPhotoMeasure' label='치수측정' size='large' getImage='GetIcon' onAction='OpenMeasurement' screentip='스케일바로 보정하고 사진의 치수 측정' supertip='사진 한 장을 선택하세요. 수동 선·평행선·3점원으로 스케일을 보정하고 거리·각도·면적을 측정합니다. 측정한 사진과 결과만 현재 슬라이드에 복사합니다.'/>
                    <button id='labMeasurementReset' label='측정 초기화' imageMso='ResetPicture' onAction='ResetMeasurement' supertip='측정 도형과 스케일이 없는 사진 복사본을 현재 슬라이드에 만듭니다.'/>
                  </group>
                <group id='labPhotoHelpGroup' label='도움말'>
                <button id='labPhotoHelp' label='사용 안내' imageMso='Help' onAction='OpenHelp'/>
                </group>
                </tab></tabs></ribbon></customUI>");
            return xml.ToString();
        }

        public object GetIcon(object control) { return RibbonIcons.Get((string)((dynamic)control).Id); }
        public void OpenImages(object control) { ShowTool("image"); }
        public void OpenBackground(object control) { ShowTool("background"); }
        public void OpenLayout(object control) { ShowTool("layout"); }
        public void OpenSpacing(object control) { ShowTool("spacing"); }
        public void OpenMeasurement(object control)
        {
            Run(delegate
            {
                if(currentForm!=null){currentForm.Activate();return;}
                PowerPointHost host=new PowerPointHost(application);SelectionSnapshot selection=host.ReadMeasurementSelection();
                try
                {
                    using(MeasurementForm form=new MeasurementForm(host,selection))
                    {currentForm=form;if(host.WindowHandle==IntPtr.Zero)form.ShowDialog();else form.ShowDialog(new WindowOwner(host.WindowHandle));}
                }
                finally{currentForm=null;}
            });
        }
        public void ResetMeasurement(object control)
        {Run(delegate{PowerPointHost host=new PowerPointHost(application);host.ResetMeasurements(host.ReadMeasurementSelection());});}
        public void ArrangeSmart(object control)
        {
            Run(delegate { new PowerPointHost(application).ArrangeAllPhotos(); });
        }
        public void AddNumber(object control)
        {
            Run(delegate
            {
                string[] tag = ((string)((dynamic)control).Tag).Split(':');
                int number;
                if (tag.Length != 2 || !NumberLabels.IsStyle(tag[0]) || (tag[1] != "next" && !int.TryParse(tag[1], out number)))
                    throw new InvalidOperationException("번호 형식을 확인할 수 없습니다.");
                new PowerPointHost(application).AddNumberLabel(tag[0], tag[1] == "next" ? (int?)null : int.Parse(tag[1]));
            });
        }
        private void Run(Action action)
        {
            try { action(); }
            catch (Exception ex) { MessageBox.Show(ex.Message, "Lab Photo Tools", MessageBoxButtons.OK, MessageBoxIcon.Information); }
        }
        public void OpenHelp(object control)
        {
            MessageBox.Show("Lab Photo Tools의 치수측정: 사진 한 장을 선택한 뒤 자 모양 버튼을 누릅니다. 스케일바 기준을 지정하고 실제 길이·단위를 입력하여 스케일을 적용하세요. 선·원·사각형·타원·각도·면적을 측정하고 측정한 사진과 결과만 현재 슬라이드에 복사합니다. n점원은 경계의 여러 점을 부드럽게 잇고 면적·둘레를 표시합니다. 측정값을 선택한 뒤 점편집에서 점과 곡률 핸들을 드래그할 수 있습니다. 그룹 사진도 사진과 측정 결과만 복사하며 기존 번호 라벨은 원본에 남습니다.\n\n사진 회전: 선택한 사진을 0.1° 단위로 회전하고 가장자리를 자릅니다.\n배경지우기: 별도 창에서 배경 제거 결과를 확인합니다. 두 기능 모두 원본 슬라이드를 복제하여 적용합니다.\n\n자동 배열과 간격 조절: 사진·도형·텍스트를 함께 선택해도 선택된 사진만 처리합니다.\n알아서 배열: 사진이 선택되어 있으면 그 사진만, 선택이 없으면 현재 슬라이드의 사진을 비율에 맞춰 앞 행부터 채웁니다. 4~16장은 정해진 촘촘한 격자로 배열하며 모든 행의 왼쪽을 맞추고 중앙 80% 영역을 사용합니다. 라벨 사진은 라벨과 사진의 그룹 전체를 함께 이동·크기 조절하여 왼쪽 위 정렬을 유지합니다.\n\n빠른 번호 매기기: 리본의 숫자나 알파벳을 누르면 현재 상태에서 라벨이 없는 사진을 왼쪽 위부터 순서대로 찾아 사진의 왼쪽 위에 붙이고 사진과 그룹화합니다. 기존 라벨을 지운 사진은 다시 라벨 대상이 됩니다. 사진이 없거나 모두 라벨이 붙었으면 독립 라벨을 빈 위치에 추가합니다. (알파벳)과 알파벳)은 A~K를 바로 표시하며 각 형식의 ‘L+’는 다음 알파벳을 추가합니다. 숫자 형식의 ‘11+’는 같은 형식의 가장 큰 번호 다음 값(최소 11)을 추가합니다. 리본 오른쪽에서 글꼴·크기·글자 색과 상자 테두리 색·굵기·바탕색을 바로 바꾸세요. 상자 색에서 ‘기본’은 형식의 기본 서식, ‘없음’은 투명입니다. 입력한 크기나 색상 코드는 Enter로 확정합니다. 설정은 자동으로 저장되어 이후 만드는 모든 번호에 적용됩니다. 17 pt 라벨의 기본 상자 한 변·원 지름은 0.85 cm이며 글자 크기에 비례합니다. 긴 번호는 글자가 잘리지 않도록 필요한 만큼 커집니다.\n\n치수측정은 사진을 왼쪽에 크게 표시하고 오른쪽 설정을 스크롤하여 모두 볼 수 있습니다. 배치와 번호 추가는 Ctrl+Z로 취소할 수 있습니다.\n사진 처리는 이 PC 안에서 실행됩니다.", "Lab Photo Tools 0.1.17", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        private void ShowTool(string mode)
        {
            Run(delegate
            {
                if (currentForm != null) { currentForm.Activate(); return; }
                PowerPointHost host = new PowerPointHost(application);
                SelectionSnapshot selection = host.ReadSelection();
                try
                {
                    using (ToolForm form = new ToolForm(host, selection, mode))
                    {
                        currentForm = form;
                        if (host.WindowHandle == IntPtr.Zero) form.ShowDialog();
                        else form.ShowDialog(new WindowOwner(host.WindowHandle));
                    }
                }
                finally { currentForm = null; }
            });
        }
    }
}
