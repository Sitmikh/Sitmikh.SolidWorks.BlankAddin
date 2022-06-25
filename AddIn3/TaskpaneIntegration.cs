using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swpublished;

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

using System.Diagnostics;

namespace Sitmikh.SolidWorks.BlankAddin
{
    /// <summary>
    /// Интеграция панели задач Add-in (делаем видимым для COM и создаем для добавления свой GUID
    /// </summary>

    [System.Runtime.InteropServices.ComVisible(true), Guid("35219AE7-988F-4181-B977-A5D36416CAE3")]
    public class TaskpaneIntegration : ISwAddin
    {
        #region Private Members

        /// <summary>
        /// Файл cookie - это текущий экземпляр SW, внутри которого мы храним ID
        /// </summary>
        private int mSwCookie;

        /// <summary>
        /// Вид панели задач для Add-in
        /// </summary>
        private TaskpaneView mTaskpaneView;

        /// <summary>
        /// Элемент управления пользовательского интерфейса, который будет находиться внутри SW taskpane view.
        /// </summary>
        private TaskpaneHostUI mTaskpaneHost;

        /// <summary>
        /// Текущий экземпляр приложения SW
        /// </summary>
        public static SldWorks mSolidWorksApplication; //!! изменил на паблик статик

        #endregion

        #region Public Members

        /// <summary>
        /// Уникальный идентификатор используемый панелью задач для регистрации в COM
        /// </summary>
        public const string SWTASKPANE_PROGID = "Sitmikh.SolidWorks.BlankAddin.Taskpane";

        #endregion

        /// <summary>
        /// Вызывается, когда SW загрузил наш Add-in и хочет, чтобы мы выполнили нашу логику подключения (вызываем не мы, а сам SW? когда открывается и закрывается)
        /// </summary>
        /// <param name="ThisSW">Текущий экземпляр SW</param>
        /// <param name="Cookie">Текущий SW cookie Id</param>
        /// <returns></returns>
        #region Solidworks Add-in Callbacks
        public bool ConnectToSW(object ThisSW, int Cookie)
        {
            // Сохранить ссылку на текущий экземпляр SW
            mSolidWorksApplication = (SldWorks)ThisSW;

            // Сохранить cookie ID
            mSwCookie = Cookie;

            // Настройка обратной связи
            var ok = mSolidWorksApplication.SetAddinCallbackInfo2(0, this, mSwCookie);

            // Создаёт наш UI
            LoadUI();

            // возврат ok
            return true;
        }


        /// <summary>
        /// Вызывается, когда SW загрузил наш Add-in и хочет, чтобы мы выполнили нашу логику подключения (вызываем не мы, а сам SW? когда открывается и закрывается)
        /// </summary>
        /// <returns></returns>
        public bool DisconnectFromSW()
        {
            // Clean up our UI
            UnloadUI();

            // Return ok
            return true;
        }
        #endregion

        #region Create UI
        /// <summary>
        /// Создайте свою панель задач и внедрите наш хост-интерфейс
        /// </summary>
        private void LoadUI()
        {
            // Найти местоположение нашего значка для панели задач (+убираем лишнее, чтобы путь был в классическом стиле)
            var imagePath = Path.Combine(Path.GetDirectoryName(typeof(TaskpaneIntegration).Assembly.CodeBase).Replace(@"file:\", string.Empty), "SmallLogo.png");

            //Создание пользовательского интерфейса в панели задач
            mTaskpaneView = mSolidWorksApplication.CreateTaskpaneView2(imagePath, "Clutch Add-in");

            //Загрузика пользовательского интерфейса в панель задач (находит класс, находящийся в dll с этим ID, чтобы внедрить его в UI)
            mTaskpaneHost = (TaskpaneHostUI)mTaskpaneView.AddControl(TaskpaneIntegration.SWTASKPANE_PROGID, string.Empty);

        }

        /// <summary>
        /// Очистить панель задач при отключении/выгрузке
        /// </summary>
        private void UnloadUI()
        {
            mTaskpaneHost = null;

            // Удалить вид панели задач
            mTaskpaneView.DeleteView();

            // Освобождаем ссылку COM и очищаем память
            Marshal.ReleaseComObject(mTaskpaneView);

            mTaskpaneView = null;
        }
        #endregion

        #region COM Registration

        /// <summary>
        /// Вызов регистрации COM для добавления записей реестра в реестр SW (работать будет просто с атрибутом. Дальнейший код для того, чтобы мы нашли добавление в настройках SW) 
        /// </summary>
        /// <param name="t"></param>
        [ComRegisterFunction()]
        private static void ComRegister(Type t)
        {
            var keyPath = string.Format(@"SOFTWARE\SolidWorks\AddIns\{0:b}", t.GUID);

            //Создание папки для регистрации добавления
            using (var rk = Microsoft.Win32.Registry.LocalMachine.CreateSubKey(keyPath))
            {
                //Загрузка добавления когда открывается SW
                rk.SetValue(null, 1);

                // Установить заголовок и описание добавления SW
                rk.SetValue("Title", "Clutch Add-in");
                rk.SetValue("Description", "SW Clutch Add-in!");
            }
        }

        /// <summary>
        /// Вызов отмены регистрации COM для удаления пользовательских записей, которые были добавлены в функцию регистрации COM
        /// </summary>
        /// <param name="t"></param>
        [ComUnregisterFunction()]
        private static void ComUnregister(Type t)
        {
            var keyPath = string.Format(@"SOFTWARE\SolidWorks\AddIns\{0:b}", t.GUID);

            // Удалить нашу запись в реестре
            Microsoft.Win32.Registry.LocalMachine.DeleteSubKeyTree(keyPath);

        }

        #endregion
    }
}
