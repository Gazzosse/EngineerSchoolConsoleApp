using DocsVision.BackOffice.CardLib.CardDefs;
using DocsVision.BackOffice.ObjectModel.Services;
using DocsVision.BackOffice.ObjectModel;
using DocsVision.Platform.ObjectManager;
using DocsVision.Platform.ObjectModel;
using DocsVision.Platform.ObjectModel.Search;
using DocumentFormat.OpenXml.Office.CustomUI;
using DocsVision.Platform.ObjectManager.Metadata;
using static DocsVision.BackOffice.CardLib.CardDefs.CardDocument;

namespace EngineerSchoolConsoleApp
{
    internal class Program
    {
        static void Main()
        {
            var serverURL = System.Configuration.ConfigurationManager.AppSettings["DVUrl"];
            var username = System.Configuration.ConfigurationManager.AppSettings["Username"];
            var password = System.Configuration.ConfigurationManager.AppSettings["Password"];

            var sessionManager = SessionManager.CreateInstance();
            sessionManager.Connect(serverURL, String.Empty, username, password);

            UserSession? session = null;
            try
            {
                session = sessionManager.CreateSession();
                var context = CreateContext(session);
                CreateApplicationBusinessTripCard(session, context);
                Console.WriteLine("Press any key to continue...");
                Console.ReadKey();
            }
            finally
            {
                session?.Close();
            }
        }

        public static ObjectContext CreateContext(UserSession session)
        {
            return DocsVision.BackOffice.ObjectModel.ContextFactory.CreateContext(session);
        }

        static void ChangeCardState(ObjectContext context, Document card, string targetState)
        {
            IStateService stateSvc = context.GetService<IStateService>();
            var branch = stateSvc.FindLineBranchesByStartState(card.SystemInfo.State)
                .FirstOrDefault(s => s.EndState.LocalizedName == targetState);
            stateSvc.ChangeState(card, branch);
        }

        public static void CreateApplicationBusinessTripCard(UserSession session, ObjectContext context)
        {
            Console.WriteLine($"Session: {session.Id}");

            var applicationBusinessTripKind = context.FindObject<KindsCardKind>(
                new QueryObject
                (
                    KindsCardKind.NameProperty.Name, "Заявка на командировку"));

            var documentService = context.GetService<IDocumentService>();
            var staffService = context.GetService<IStaffService>();
            var partnersService = context.GetService<IPartnersService>();
            var baseUniversalService = context.GetService<IBaseUniversalService>();
            var fileCardService = context.GetService<IVersionedFileCardService>();

            var cityType = baseUniversalService.FindItemTypeWithSameName("Города", null);
            var cityItem = baseUniversalService.FindItemWithSameName("Санкт-Петербург", cityType);


            var applicationBusinessTrip = documentService.CreateDocument(null, applicationBusinessTripKind);

            applicationBusinessTrip.MainInfo.Author = staffService.GetCurrentEmployee();
            applicationBusinessTrip.MainInfo.Name = "Код5";
            applicationBusinessTrip.MainInfo["RegDate"] = DateTime.Now;
            applicationBusinessTrip.MainInfo.Registrar = staffService.FindEmpoyeeByAccountName("ENGINEER\\ivanov_as");
            applicationBusinessTrip.MainInfo["BusinessTripStart"] = new DateTime(2025, 11, 3);
            applicationBusinessTrip.MainInfo["BusinessTripEnd"] = new DateTime(2025, 11, 4);
            applicationBusinessTrip.MainInfo["Cities"] = cityItem.GetObjectId();
            applicationBusinessTrip.MainInfo["BusinessTripDuration"] = 2;
            applicationBusinessTrip.MainInfo["BusinessTripExpenses"] = 1000m;
            applicationBusinessTrip.MainInfo["Organization"] = partnersService.FindCompanyByNameOnServer(null, "Галиулин Ко. Холдингс").GetObjectId();
            applicationBusinessTrip.MainInfo["BusinessTripReason"] = "Самое обоснованное обоснование для поездки";

            var arrangers = (IList<BaseCardSectionRow>)applicationBusinessTrip.GetSection(new Guid("D8A59AFE-3118-4AEB-9419-59DE69D4B622"));
            var arrangerRow = new BaseCardSectionRow();
            arrangerRow["Arranger"] = staffService.FindEmpoyeeByAccountName("ENGINEER\\kolesnikova_sn").GetObjectId();
            arrangers.Add(arrangerRow);

            var approvers = (IList<BaseCardSectionRow>)applicationBusinessTrip.GetSection(CardDocument.Approvers.ID);
            var approverRow = new BaseCardSectionRow();
            approverRow["Approver"] = staffService.FindEmpoyeeByAccountName("ENGINEER\\mikhailov_sa").GetObjectId();
            approvers.Add(approverRow);

            applicationBusinessTrip.MainInfo["Tickets"] = 1;

            var secondedEmployee = staffService.FindEmpoyeeByAccountName("ENGINEER\\samoilov_pn");
            applicationBusinessTrip.MainInfo["SecondedEmployee"] = secondedEmployee.GetObjectId();
            applicationBusinessTrip.MainInfo["Manager"] = secondedEmployee.Unit.Manager.GetObjectId();
            applicationBusinessTrip.MainInfo["WorkPhoneNumber"] = secondedEmployee.Unit.Manager.Phone;

            context.AcceptChanges();

            /*FileData file = session.FileManager.CreateFile("ИПР_I_SLAE");
            file.Upload("C:\\StudyV2\\EngineerSchoolConsoleApp\\Files\\ИПР_I_SLAE.pdf");
            var files = (IList<BaseCardSectionRow>)applicationBusinessTrip.GetSection(CardDocument.Files.ID);
            var fileRow = new BaseCardSectionRow();
            fileRow[CardDocument.Files.FileId] = file.Id;
            files.Add(fileRow);*/

            var fileCard = fileCardService.CreateCard(@"C:\StudyV2\EngineerSchoolConsoleApp\Files\ИПР_I_SLAE.pdf");

            var files = (IList<BaseCardSectionRow>)applicationBusinessTrip.GetSection(CardDocument.Files.ID);
            var fileRow = new BaseCardSectionRow();
            fileRow[CardDocument.Files.FileId] = fileCard.Id;
            files.Add(fileRow);

            context.AcceptChanges();

            ChangeCardState(context, applicationBusinessTrip, "На согласовании");
            context.AcceptChanges();

            Console.WriteLine($"New card id: {applicationBusinessTrip.GetObjectId()}");
        }

        public static void SomeLogic(UserSession session, ObjectContext context)
        {
            Console.WriteLine($"Session: {session.Id}");

            var officeMemoKind = context.FindObject<KindsCardKind>(
                new QueryObject(
                    KindsCardKind.NameProperty.Name, "Служебная записка"));

            var requestTypeId = new Guid("{12A19587-C6C0-477F-9811-EFEBAB3FBBE3}");
            var requestType = context.GetObject<BaseUniversalItem>(requestTypeId);

            var docSvc = context.GetService<IDocumentService>();
            var staffSvc = context.GetService<IStaffService>();
            var officeMemo = docSvc.CreateDocument(null, officeMemoKind);
            officeMemo.MainInfo.Author = staffSvc.GetCurrentEmployee();
            officeMemo.MainInfo.Registrar = staffSvc.GetCurrentEmployee();
            officeMemo.MainInfo[CardDocument.MainInfo.RegDate] = DateTime.Now;
            officeMemo.MainInfo.Name = "Card created from code";
            officeMemo.MainInfo.Item = requestType;

            var approvers = (IList<BaseCardSectionRow>)officeMemo.GetSection(CardDocument.Approvers.ID);
            var approverRow1 = new BaseCardSectionRow();
            approverRow1[CardDocument.Approvers.Approver] = staffSvc.GetCurrentEmployee().GetObjectId();
            approvers.Add(approverRow1);
            var approverRow2 = new BaseCardSectionRow();
            approverRow2[CardDocument.Approvers.Approver] = staffSvc.FindEmpoyeeByAccountName("ENGINEER\\DVAdmin")?.GetObjectId();
            approvers.Add(approverRow2);

            context.AcceptChanges();

            ChangeCardState(context, officeMemo, "Is approving");
            context.AcceptChanges();

            Console.WriteLine($"New card id: {officeMemo.GetObjectId()}");
        }
    }
}
