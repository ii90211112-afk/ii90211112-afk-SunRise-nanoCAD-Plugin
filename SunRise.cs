using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using HostMgd.ApplicationServices; // Для HostApplication
using System.Management;
using Teigha.Runtime;
using System.Windows.Forms;
using System.Globalization;

// 1. Обязательный атрибут, который указывает nanoCAD на ваш стартовый класс
[assembly: ExtensionApplication(typeof(SunRise.SunRise))]

namespace SunRise
{
    
    using System.Globalization;
    using System.Reflection;
    using System.Runtime.InteropServices;
    using System.Text.RegularExpressions;
    using HostMgd.EditorInput;
    using HostMgd.Windows;
    using Teigha.DatabaseServices;
    using Teigha.LayerManager;
    using Teigha.Runtime;

    public class SunRise: IExtensionApplication
    {
        public void Initialize()
        {
            var ed = HostMgd.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
            ed.WriteMessage("Расширение SunRise v:{0} успешно загружено", Assembly.GetExecutingAssembly().GetName().Version);
            ed.WriteMessage("Введите команду SunRiseInfo, для получения подробностей");
        }

        public void Terminate()
        {
            // Здесь можно освобождать ресурсы при закрытии nanoCAD
        }

        [CommandMethod("SunRiseInfo")]
        public void Info()
        {
            var ed = HostMgd.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
            ed.WriteMessage("Расширение SunRise версия " + Assembly.GetExecutingAssembly().GetName().Version); 
            ed.WriteMessage("Доступные команды: ");
            ed.WriteMessage("CheckDisk - проверяет СМАРТ-статуса диска, может не работать без админских прав");
            ed.WriteMessage("PaintItBlack - красит все объекты в черный цвет, даже внутри блоков может быть необходимо для печати. Примечание: чтобы перекрасить МТекст-ы, нужно их разбить");
            ed.WriteMessage("ConvertToMText - преобразует текст(или тексты) в МТекст. При указании более одного текста - объединяет их в один МТекст");
            ed.WriteMessage("FixTrueColors - преобразует цвета из х,х,х в цвета NanoCad - особо часто требуется для печати");
            ed.WriteMessage("FlattenDrawing - Сплющивает весь чертеж. Задает z=0 для всех объектов в т.ч в блоках. Также задает z=0 для всех точек сплайнов и т.д");
            ed.WriteMessage("SR_SHOW_BLOCK_STRUCTURE - выписывает в консоль содержимое блока с указанием слоев. Вложенные блоки не рассматриваются");
            ed.WriteMessage("SR_FIND_LAYER_IN_BLOCKS - выписывает в консоль блоки, которые используют указанный пользователем слой внутри себя");
            ed.WriteMessage("SR_CLEAN_LAYER_IN_BLOCK - переносит все блоки с указанного пользователем слоя на слой \"0\" ");

        }

        /// <summary>
        /// Команда для проверки СМАРТ-статуса диска, может не работать без админских прав
        /// </summary>
        [CommandMethod ("CheckDisk")]
        public static void CheckStatus()
        {
            var ed = HostMgd.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
            try
            {
                // Запрос к WMI для получения модели и статуса всех дисков
                ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT Model, Status FROM Win32_DiskDrive");

                foreach (ManagementObject drive in searcher.Get())
                {
                    string model = drive["Model"]?.ToString();
                    string status = drive["Status"]?.ToString();

                    ed.WriteMessage("Диск: {0}", model);
                    ed.WriteMessage("S.M.A.R.T. Статус: {0}", status);
                    ed.WriteMessage("------------------------------");
                }
            }
            catch (Exception ex)
            {
                ed.WriteMessage("Ошибка при получении данных: " + ex.Message);
            }
        }

        /// <summary>
        /// Копия PaintItBlack2 из ClassLibrary1
        /// красит все объекты на чертеже в черный цвет, даже внутри блоков, может быть необходимо для печати
        /// Примечание: чтобы перекрасить МТекст-ы, нужно их разбить
        /// </summary>
        [CommandMethod("PaintItBlack")]
        public void PaintItBlack()
        {
            var ed = HostMgd.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
            try
            {
                // 1. Подключаемся к COM
                object comApp = System.Runtime.InteropServices.Marshal.GetActiveObject("nanoCAD.Application.5.0");
                object comDoc = comApp.GetType().InvokeMember("ActiveDocument", BindingFlags.GetProperty, null, comApp, null);

                // 2. Получаем коллекцию БЛОКОВ (это описания того, что внутри)
                object blocks = comDoc.GetType().InvokeMember("Blocks", BindingFlags.GetProperty, null, comDoc, null);
                int blocksCount = (int)blocks.GetType().InvokeMember("Count", BindingFlags.GetProperty, null, blocks, null);

                ed.WriteMessage("\n--- Обработка определений блоков (" + blocksCount + ") ---");
                int entitiesChanged = 0;

                // Перебираем все определения блоков в чертеже
                for (int i = 0; i < blocksCount; i++)
                {
                    object block = blocks.GetType().InvokeMember("Item", BindingFlags.InvokeMethod, null, blocks, new object[] { i });
                    int entCount = (int)block.GetType().InvokeMember("Count", BindingFlags.GetProperty, null, block, null);

                    // Перебираем объекты ВНУТРИ каждого блока
                    for (int j = 0; j < entCount; j++)
                    {
                        try
                        {
                            object entity = block.GetType().InvokeMember("Item", BindingFlags.InvokeMethod, null, block, new object[] { j });

                            // Устанавливаем цвет 7 (Черный/Белый). 
                            // Можно также поставить 0 (ByBlock), чтобы они слушались цвета самого блока.
                            entity.GetType().InvokeMember("Color", BindingFlags.SetProperty, null, entity, new object[] { 7 });
                            entitiesChanged++;
                        }
                        catch { continue; }
                    }
                }

                // 3. Также пройдемся по объектам в Модели (на случай, если там есть не блоки)
                object modelSpace = comDoc.GetType().InvokeMember("ModelSpace", BindingFlags.GetProperty, null, comDoc, null);
                int modelCount = (int)modelSpace.GetType().InvokeMember("Count", BindingFlags.GetProperty, null, modelSpace, null);
                for (int k = 0; k < modelCount; k++)
                {
                    try
                    {
                        object entity = modelSpace.GetType().InvokeMember("Item", BindingFlags.InvokeMethod, null, modelSpace, new object[] { k });
                        entity.GetType().InvokeMember("Color", BindingFlags.SetProperty, null, entity, new object[] { 7 });
                        if (entity.GetType() == typeof(MText)) 
                        {
                            
                        }
                    }
                    catch { continue; }
                }

                // Обновляем экран
                comDoc.GetType().InvokeMember("Regen", BindingFlags.InvokeMethod, null, comDoc, new object[] { 1 });
                ed.WriteMessage("\n[Успех] Все объекты внутри блоков перекрашены. Изменено элементов: " + entitiesChanged);
            }
            catch (Exception ex)
            {
                ed.WriteMessage("\n[Ошибка]: " + (ex.InnerException?.Message ?? ex.Message));
            }
        }

        /// <summary>
        /// ConvertToMtext преобразует один или больше DBText в одиин MText
        /// </summary>
        [CommandMethod("ConvertToMText")]
        public void ToUniteTexts()
        {
            List<DBText> selectedDBTexts = new List<DBText>();
            List<string> uniteTexts = new List<string>();
            DBText text = new DBText();//чтобы передать шрифт

            Document doc = HostMgd.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            //выбор объектов
            PromptSelectionResult selResult = ed.GetSelection();
            SelectionSet selectionSet = selResult.Value;
            // Получаем Transaction для работы с БД
            using (Transaction tr = doc.TransactionManager.StartTransaction())
            {
                foreach (SelectedObject selObj in selectionSet)
                {
                    if (selObj != null)
                    {
                        DBText ent = tr.GetObject(selObj.ObjectId, OpenMode.ForRead) as DBText;

                        if (ent != null)
                        {
                            selectedDBTexts.Add(ent);
                        }
                    }
                }

                foreach (DBText dBText in selectedDBTexts)
                {
                    if (dBText.GetType() == typeof(DBText))
                    {
                        text = (DBText)dBText;//только чтобы сохранить шрифт
                        uniteTexts.Add(dBText.TextString);
                    }
                }

                //Создание МТекст-а
                MText mText = new MText();
                mText.SetPropertiesFrom(text);
                mText.TextHeight = text.Height;
                mText.Location = selectedDBTexts[0].Position;

                foreach (string str in uniteTexts)
                {
                    mText.Contents += str + "\n";
                }

                //Добавление на чертеж
                BlockTable blockTable = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead, false);
                BlockTableRecord record = (BlockTableRecord)tr.GetObject(blockTable[BlockTableRecord.ModelSpace], OpenMode.ForWrite, false);
                record.AppendEntity(mText);
                tr.AddNewlyCreatedDBObject(mText, true);

                tr.Commit();
            }


        }

        /// <summary>
        /// Переписать цвета
        /// </summary>
        [CommandMethod("FixTrueColors")]
        public void FixTrueColors()
        {
            var doc = HostMgd.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;

            var db = doc.Database;
            var ed = doc.Editor;

            int fixedLayersCount = 0;
            int fixedObjectsCount = 0;
            int fixedMTextCount = 0;

            using (var tr = db.TransactionManager.StartTransaction())
            {
                // 1. ИСПРАВЛЯЕМ ВСЕ СЛОИ В ЧЕРТЕЖЕ
                var lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
                foreach (ObjectId layerId in lt)
                {
                    var layer = (LayerTableRecord)tr.GetObject(layerId, OpenMode.ForRead);
                    if (layer.Color.IsByColor)
                    {
                        layer.UpgradeOpen();
                        var indexColor = Teigha.Colors.Color.FromRgb(layer.Color.Red, layer.Color.Green, layer.Color.Blue);
                        layer.Color = Teigha.Colors.Color.FromColorIndex(Teigha.Colors.ColorMethod.ByAci, indexColor.ColorIndex);
                        fixedLayersCount++;
                    }
                }

                // 2. РЕКУРСИВНЫЙ ОБХОД ВСЕХ БЛОКОВ И ПРОСТРАНСТВ
                // BlockTable содержит описания всех блоков чертежа, включая пространства Модели и Листов.
                // Проходя по нему, мы автоматически чистим объекты "внутри" любых блоков, независимо от уровня вложенности.
                var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                foreach (ObjectId btrId in bt)
                {
                    var btr = (BlockTableRecord)tr.GetObject(btrId, OpenMode.ForRead);

                    // Игнорируем анонимные блоки динамических свойств, если это необходимо,
                    // но для тотальной очистки цветов лучше пройтись по всем без исключения.
                    foreach (ObjectId entId in btr)
                    {
                        var ent = tr.GetObject(entId, OpenMode.ForRead) as Entity;
                        if (ent == null) continue;

                        // А) Исправляем стандартное свойство цвета True Color
                        if (ent.Color.IsByColor)
                        {
                            ent.UpgradeOpen();
                            var indexColor = Teigha.Colors.Color.FromRgb(ent.Color.Red, ent.Color.Green, ent.Color.Blue);
                            ent.Color = Teigha.Colors.Color.FromColorIndex(Teigha.Colors.ColorMethod.ByAci, indexColor.ColorIndex);
                            fixedObjectsCount++;
                        }

                        // Б) Лечим скрытый RGB-цвет внутри форматирования MText
                        if (ent is MText mtext)
                        {
                            // Проверяем, содержит ли текст управляющие коды цветов True Color (формат \C... или \c...)
                            // В САПР True Color в коде текста часто пишется как большое целое число (например, \C16777216;)
                            if (mtext.Contents.Contains("\\C") || mtext.Contents.Contains("\\c"))
                            {
                                ent.UpgradeOpen();

                                // Регулярное выражение ищет теги цвета типа \C12345678; или \c123456;
                                // и удаляет их, возвращая текст к цвету "ПоСлою" (ByLayer)
                                string updatedContents = Regex.Replace(mtext.Contents, @"\\[Cc][0-9]+;", "");

                                // Если текст изменился, перезаписываем его свойства
                                if (mtext.Contents != updatedContents)
                                {
                                    mtext.Contents = updatedContents;
                                    fixedMTextCount++;
                                }
                            }
                        }

                        // В) для таблиц - тесты показали что это не работает, но мало-ли
                        if (ent is Table table)
                        {
                            table.UpgradeOpen();

                            // ПОЯЧЕЕЧНЫЙ ПЕРЕБОР — самый надежный способ для nanoCAD 5.1
                            for (int row = 0; row < table.NumRows; row++)
                            {
                                for (int col = 0; col < table.NumColumns; col++)
                                {
                                    // 1. Исправляем цвет текста в ячейке
                                    var contentColor = table.GetContentColor(row, col, 0);
                                    if (contentColor.IsByColor)
                                    {
                                        var indexColor = Teigha.Colors.Color.FromRgb(contentColor.Red, contentColor.Green, contentColor.Blue);
                                        table.SetContentColor(row, col, 0, Teigha.Colors.Color.FromColorIndex(Teigha.Colors.ColorMethod.ByAci, indexColor.ColorIndex));
                                        fixedObjectsCount++;
                                    }

                                    // 2. Исправляем цвет индивидуальных границ ячейки
                                    // Перебираем значения от 1 до 63, преобразуя их в GridLineType.
                                    // Так как компилятор не ругался на эту сигнатуру на прошлом шаге — она верная!
                                    for (int gridInt = 1; gridInt <= 63; gridInt++)
                                    {
                                        var currentGridType = (Teigha.DatabaseServices.GridLineType)gridInt;
                                        try
                                        {
                                            var gridColor = table.GetGridColor(row, col, currentGridType);
                                            if (gridColor.IsByColor)
                                            {
                                                var indexColor = Teigha.Colors.Color.FromRgb(gridColor.Red, gridColor.Green, gridColor.Blue);

                                                // Используем сигнатуру SetGridColor для ячеек, которая точно совпадает с GetGridColor
                                                table.SetGridColor(row, col, currentGridType, Teigha.Colors.Color.FromColorIndex(Teigha.Colors.ColorMethod.ByAci, indexColor.ColorIndex));
                                                fixedObjectsCount++;
                                            }
                                        }
                                        catch
                                        {
                                            // Игнорируем ошибки для неподдерживаемых комбинаций флагов
                                        }
                                    }
                                }
                            }
                        }


                    }
                }

                tr.Commit();
            }

            // Выводим красивый детальный отчет в консоль nanoCAD
            ed.WriteMessage($"\n[SunRise]: Глубокая очистка True Color завершена успешно!");
            ed.WriteMessage($"\n - Слоев переведено из RGB: {fixedLayersCount}");
            ed.WriteMessage($"\n - Графических объектов исправлено: {fixedObjectsCount}");
            if (fixedMTextCount > 0)
            {
                ed.WriteMessage($"\n - Очищено скрытых RGB-тегов внутри МТекстов: {fixedMTextCount}");
            }
            ed.Regen();
        }


        /// <summary>
        /// Делает весь чертеж плоским
        /// </summary>
        [CommandMethod("FlattenDrawing")]
        public void FlattenDrawing()
        {
            var doc = HostMgd.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;

            var db = doc.Database;
            var ed = doc.Editor;

            int flattenedCount = 0;

            using (var tr = db.TransactionManager.StartTransaction())
            {
                // Получаем таблицу блоков (она хранит пространства Модели, Листов и описания всех блоков)
                var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);

                foreach (ObjectId btrId in bt)
                {
                    var btr = (BlockTableRecord)tr.GetObject(btrId, OpenMode.ForRead);

                    foreach (ObjectId entId in btr)
                    {
                        var ent = tr.GetObject(entId, OpenMode.ForRead) as Entity;
                        if (ent == null) continue;

                        bool isModified = false;

                        // 1. ОТРЕЗКИ (Line)
                        if (ent is Line line)
                        {
                            if (line.StartPoint.Z != 0 || line.EndPoint.Z != 0)
                            {
                                line.UpgradeOpen();
                                line.StartPoint = new Teigha.Geometry.Point3d(line.StartPoint.X, line.StartPoint.Y, 0);
                                line.EndPoint = new Teigha.Geometry.Point3d(line.EndPoint.X, line.EndPoint.Y, 0);
                                isModified = true;
                            }
                        }
                        // 2. ОБЫЧНЫЕ ПОЛИЛИНИИ (Polyline / LwPolyline)
                        else if (ent is Polyline poly)
                        {
                            if (poly.Elevation != 0)
                            {
                                poly.UpgradeOpen();
                                poly.Elevation = 0;
                                isModified = true;
                            }
                        }
                        // 3. ТЕКСТЫ (DBText)
                        else if (ent is DBText text)
                        {
                            if (text.Position.Z != 0 || text.AlignmentPoint.Z != 0)
                            {
                                text.UpgradeOpen();
                                text.Position = new Teigha.Geometry.Point3d(text.Position.X, text.Position.Y, 0);
                                text.AlignmentPoint = new Teigha.Geometry.Point3d(text.AlignmentPoint.X, text.AlignmentPoint.Y, 0);
                                isModified = true;
                            }
                        }
                        // 4. МУЛЬТИТЕКСТЫ (MText)
                        else if (ent is MText mtext)
                        {
                            if (mtext.Location.Z != 0)
                            {
                                mtext.UpgradeOpen();
                                mtext.Location = new Teigha.Geometry.Point3d(mtext.Location.X, mtext.Location.Y, 0);
                                isModified = true;
                            }
                        }
                        // 5. КРУГИ (Circle)
                        else if (ent is Circle circle)
                        {
                            if (circle.Center.Z != 0 || circle.Normal.Z != 1)
                            {
                                circle.UpgradeOpen();
                                circle.Center = new Teigha.Geometry.Point3d(circle.Center.X, circle.Center.Y, 0);
                                circle.Normal = new Teigha.Geometry.Vector3d(0, 0, 1);
                                isModified = true;
                            }
                        }
                        // 6. ДУГИ (Arc)
                        else if (ent is Arc arc)
                        {
                            if (arc.Center.Z != 0 || arc.Normal.Z != 1)
                            {
                                arc.UpgradeOpen();
                                arc.Center = new Teigha.Geometry.Point3d(arc.Center.X, arc.Center.Y, 0);
                                arc.Normal = new Teigha.Geometry.Vector3d(0, 0, 1);
                                isModified = true;
                            }
                        }
                        // 7. 3D ПОЛИЛИНИИ (Polyline3d)
                        else if (ent is Polyline3d poly3d)
                        {
                            poly3d.UpgradeOpen();
                            foreach (ObjectId vertexId in poly3d)
                            {
                                var v3d = (PolylineVertex3d)tr.GetObject(vertexId, OpenMode.ForWrite);
                                v3d.Position = new Teigha.Geometry.Point3d(v3d.Position.X, v3d.Position.Y, 0);
                            }
                            isModified = true;
                        }
                        // 8. СПЛАЙНЫ (Spline)
                        else if (ent is Spline spline)
                        {
                            spline.UpgradeOpen();
                            for (int i = 0; i < spline.NumControlPoints; i++)
                            {
                                var cp = spline.GetControlPointAt(i);
                                if (cp.Z != 0)
                                {
                                    spline.SetControlPointAt(i, new Teigha.Geometry.Point3d(cp.X, cp.Y, 0));
                                    isModified = true;
                                }
                            }
                        }
                        // 9. РАЗМЕРЫ (Dimension)
                        else if (ent is Dimension dim)
                        {
                            if (dim.Normal.Z != 1 || dim.Elevation != 0)
                            {
                                dim.UpgradeOpen();
                                dim.Normal = new Teigha.Geometry.Vector3d(0, 0, 1);
                                dim.Elevation = 0;
                                isModified = true;
                            }
                        }
                        // 10. ШТРИХОВКИ (Hatch)
                        else if (ent is Hatch hatch)
                        {
                            if (hatch.Normal.Z != 1 || hatch.Elevation != 0)
                            {
                                hatch.UpgradeOpen();
                                hatch.Normal = new Teigha.Geometry.Vector3d(0, 0, 1);
                                hatch.Elevation = 0;
                                isModified = true;
                            }
                        }
                        // 11. ТОЧКИ (DBPoint) — Сдвиг через матрицу смещения
                        else if (ent is DBPoint point)
                        {
                            if (point.Position.Z != 0)
                            {
                                point.UpgradeOpen();
                                var moveVector = new Teigha.Geometry.Vector3d(0, 0, -point.Position.Z);
                                point.TransformBy(Teigha.Geometry.Matrix3d.Displacement(moveVector));
                                isModified = true;
                            }
                        }
                        // 12. ВХОЖДЕНИЯ БЛОКОВ (BlockReference) — Сдвиг + чистка вложенных атрибутов
                        else if (ent is BlockReference br)
                        {
                            if (br.Position.Z != 0)
                            {
                                br.UpgradeOpen();
                                var moveVector = new Teigha.Geometry.Vector3d(0, 0, -br.Position.Z);
                                br.TransformBy(Teigha.Geometry.Matrix3d.Displacement(moveVector));
                                isModified = true;
                            }

                            // Дополнительно лечим вложенные в блок атрибуты (AttributeReference)
                            foreach (ObjectId attId in br.AttributeCollection)
                            {
                                var att = tr.GetObject(attId, OpenMode.ForRead) as AttributeReference;
                                if (att != null && (att.Position.Z != 0 || att.AlignmentPoint.Z != 0))
                                {
                                    att.UpgradeOpen();
                                    att.Position = new Teigha.Geometry.Point3d(att.Position.X, att.Position.Y, 0);
                                    att.AlignmentPoint = new Teigha.Geometry.Point3d(att.AlignmentPoint.X, att.AlignmentPoint.Y, 0);
                                    isModified = true;
                                }
                            }
                        }
                        // 13. МУЛЬТИВЫНОСКИ (MLeader) — Сдвиг на основе вычисления её габаритов
                        else if (ent is MLeader mleader)
                        {
                            try
                            {
                                double currentZ = mleader.GeometricExtents.MinPoint.Z;
                                if (currentZ != 0)
                                {
                                    mleader.UpgradeOpen();
                                    var moveVector = new Teigha.Geometry.Vector3d(0, 0, -currentZ);
                                    mleader.TransformBy(Teigha.Geometry.Matrix3d.Displacement(moveVector));
                                    isModified = true;
                                }
                            }
                            catch
                            {
                                // На случай пустых выносок без графики
                            }
                        }

                        if (isModified)
                        {
                            flattenedCount++;
                        }
                    }
                }

                tr.Commit();
            }

            ed.WriteMessage($"\n[SunRise]: Операция завершена!");
            ed.WriteMessage($"\n - Всего приземлено на плоскость Z=0: {flattenedCount} объектов.");
            ed.Regen();
        }

        [CommandMethod("SR_SHOW_BLOCK_STRUCTURE")]
        public void ShowBlockStructureCommand()
        {
            Document doc = HostMgd.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            string blockName = string.Empty;

            // 1. Попытка интерактивного выбора блока кликом мыши
            PromptEntityOptions peo = new PromptEntityOptions("\nВыберите блок на чертеже: ");
            peo.SetRejectMessage("\nВыбранный объект не является блоком!");
            peo.AddAllowedClass(typeof(BlockReference), true); // Разрешаем выбирать только BlockReference

            PromptEntityResult per = ed.GetEntity(peo);

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                if (per.Status == PromptStatus.OK)
                {
                    // Пользователь успешно кликнул на блок
                    BlockReference blkRef = (BlockReference)tr.GetObject(per.ObjectId, OpenMode.ForRead);

                    // Получаем имя определения блока (BlockTableRecord)
                    // Использование AnonymousBlockName / DynamicBlockTableRecord корректно для динамических блоков
                    if (blkRef.IsDynamicBlock)
                    {
                        BlockTableRecord dbr = (BlockTableRecord)tr.GetObject(blkRef.DynamicBlockTableRecord, OpenMode.ForRead);
                        blockName = dbr.Name;
                    }
                    else
                    {
                        blockName = blkRef.Name;
                    }
                }
                else
                {
                    // 2. Альтернатива: если клик не удался, запрашиваем имя текстом
                    ed.WriteMessage("\nБлок не выбран. Попробуем найти по имени.");
                    PromptStringOptions pso = new PromptStringOptions("\nВведите имя блока вручную: ");
                    pso.AllowSpaces = true;
                    PromptResult pr = ed.GetString(pso);

                    if (pr.Status != PromptStatus.OK || string.IsNullOrEmpty(pr.StringResult))
                    {
                        ed.WriteMessage("\nОперация отменена.");
                        return;
                    }
                    blockName = pr.StringResult;
                }

                // 3. Анализ структуры выбранного или введенного блока
                BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);

                if (!bt.Has(blockName))
                {
                    ed.WriteMessage(string.Format("\nБлок с именем '{0}' не найден в базе данных чертежа.", blockName));
                    return;
                }

                BlockTableRecord btr = (BlockTableRecord)tr.GetObject(bt[blockName], OpenMode.ForRead);

                ed.WriteMessage(string.Format("\n\n--- Структура блока: {0} ---", blockName));

                int elementsCount = 0;
                Dictionary<string, int> stats = new Dictionary<string, int>();

                foreach (ObjectId entityId in btr)
                {
                    Entity ent = tr.GetObject(entityId, OpenMode.ForRead) as Entity;
                    if (ent != null)
                    {
                        elementsCount++;
                        string objectType = ent.GetType().Name;
                        string layerName = ent.Layer;

                        ed.WriteMessage(string.Format("\nЭлемент #{0}: Тип = {1} | Слой = {2}", elementsCount, objectType, layerName));

                        string key = string.Format("Тип: {0} на слое: {1}", objectType, layerName);
                        if (stats.ContainsKey(key))
                            stats[key]++;
                        else
                            stats[key] = 1;
                    }
                }

                if (elementsCount == 0)
                {
                    ed.WriteMessage("\nЭтот блок пуст.");
                }
                else
                {
                    ed.WriteMessage("\n\n--- Сводная статистика по слоям блока ---");
                    foreach (KeyValuePair<string, int> pair in stats)
                    {
                        ed.WriteMessage(string.Format("\n{0} — Количество: {1}", pair.Key, pair.Value));
                    }
                }

                tr.Commit();
            }
        }

        [CommandMethod("SR_FIND_LAYER_IN_BLOCKS")]
        public void FindLayerInBlocksCommand()
        {
            Document doc = HostMgd.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            // 1. Запрашиваем имя проблемного слоя
            PromptStringOptions pso = new PromptStringOptions("\nВведите имя слоя, который не удаляется: ");
            pso.AllowSpaces = true;
            PromptResult pr = ed.GetString(pso);

            if (pr.Status != PromptStatus.OK || string.IsNullOrEmpty(pr.StringResult))
            {
                ed.WriteMessage("\nОперация отменена.");
                return;
            }

            string targetLayer = pr.StringResult.Trim();

            // Проверяем, существует ли вообще такой слой в системе
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                LayerTable lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
                if (!lt.Has(targetLayer))
                {
                    ed.WriteMessage(string.Format("\nОшибка: Слой '{0}' не существует в данном чертеже.", targetLayer));
                    return;
                }

                // Вспомогательные служебные слои, которые нельзя удалить в принципе
                if (targetLayer.Equals("0", StringComparison.OrdinalIgnoreCase) ||
                    targetLayer.Equals("Defpoints", StringComparison.OrdinalIgnoreCase))
                {
                    ed.WriteMessage(string.Format("\nСлой '{0}' является системным, его невозможно удалить штатными средствами.", targetLayer));
                    return;
                }

                ed.WriteMessage(string.Format("\nНачинаю сканирование блоков на наличие объектов на слое '{0}'...", targetLayer));

                // Открываем таблицу блоков для глобального поиска
                BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);

                bool foundAny = false;
                int totalCulprits = 0;

                // 2. Обходим абсолютно все описания блоков в базе данных
                foreach (ObjectId btrId in bt)
                {
                    BlockTableRecord btr = (BlockTableRecord)tr.GetObject(btrId, OpenMode.ForRead);

                    // Пропускаем пространства листов и модели (это служебные блоки, нас интересуют только пользовательские)
                    if (btr.IsLayout)
                        continue;

                    // Список элементов внутри текущего блока, которые сидят на искомом слое
                    List<string> culpritsInBlock = new List<string>();

                    // Перебираем всю геометрию внутри описания блока
                    foreach (ObjectId entId in btr)
                    {
                        Entity ent = tr.GetObject(entId, OpenMode.ForRead) as Entity;
                        if (ent != null)
                        {
                            // Сравниваем имя слоя без учета регистра
                            if (ent.Layer.Equals(targetLayer, StringComparison.OrdinalIgnoreCase))
                            {
                                culpritsInBlock.Add(string.Format("{0} (ID: {1})", ent.GetType().Name, ent.Id.Handle.ToString()));
                                totalCulprits++;
                            }
                        }
                    }

                    // Если в блоке нашли нарушителей — выводим отчет для пользователя
                    if (culpritsInBlock.Count > 0)
                    {
                        foundAny = true;
                        ed.WriteMessage(string.Format("\n\n[Блок: {0}] содержит объектов на слое: {1}", btr.Name, culpritsInBlock.Count));
                        foreach (string info in culpritsInBlock)
                        {
                            ed.WriteMessage(string.Format("\n  -> {0}", info));
                        }
                    }
                }

                // 3. Финальный вердикт
                if (!foundAny)
                {
                    ed.WriteMessage(string.Format("\n\nСлой '{0}' НЕ используется внутри геометрии блоков.", targetLayer));
                    ed.WriteMessage("\nЕсли он всё еще не удаляется, возможные причины:");
                    ed.WriteMessage("\n - На слое висят невидимые пустые текстовые строки или точки в пространстве модели.");
                    ed.WriteMessage("\n - Слой используется в настройках размерных или текстовых стилей.");
                    ed.WriteMessage("\n - На слое лежат внешние ссылки (XRef) или пустые вставки блоков.");
                }
                else
                {
                    ed.WriteMessage(string.Format("\n\nВсего найдено объектов: {0} внутри описаний блоков чертежа.", totalCulprits));
                    ed.WriteMessage("\nРешение: Зайдите в редактор этих блоков (BEDIT) и переведите указанные объекты на слой '0' или удалите их.");
                }

                tr.Commit();
            }
        }

        [CommandMethod("SR_CLEAN_LAYER_IN_BLOCKS")]
        public void CleanLayerInBlocksCommand()
        {
            // 1. Показываем всплывающее окно с предупреждением (с явным указанием System.Windows.Forms)
            System.Windows.Forms.DialogResult result = System.Windows.Forms.MessageBox.Show(
                "Внимание! Данная команда внесет изменения в структуру блоков (включая анонимные) и перенесет объекты на слой '0'.\n\n" +
                "Рекомендуется пересохранить и сделать резервную копию файла перед продолжением.\n\n" +
                "Вы хотите продолжить?",
                "Предупреждение [SunRise]",
                System.Windows.Forms.MessageBoxButtons.YesNo,
                System.Windows.Forms.MessageBoxIcon.Warning
            );

            // Если пользователь нажал "Нет" (или закрыл окно) — прерываем выполнение
            if (result != System.Windows.Forms.DialogResult.Yes)
            {
                // Выводим сообщение в консоль nanoCAD, чтобы пользователь понял, почему ничего не произошло
                HostMgd.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage("\nОперация отменена пользователем.");
                return;
            }

            // 2. Основная логика работы (выполняется только если нажали "Да")
            Document doc = HostMgd.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            PromptStringOptions pso = new PromptStringOptions("\nВведите имя слоя для очистки (переноса на слой '0'): ");
            pso.AllowSpaces = true;
            PromptResult pr = ed.GetString(pso);

            if (pr.Status != PromptStatus.OK || string.IsNullOrEmpty(pr.StringResult)) return;
            string targetLayer = pr.StringResult.Trim();

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                LayerTable lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
                if (!lt.Has(targetLayer))
                {
                    ed.WriteMessage(string.Format("\nСлой '{0}' не найден.", targetLayer));
                    return;
                }

                BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                int changedCount = 0;

                foreach (ObjectId btrId in bt)
                {
                    BlockTableRecord btr = (BlockTableRecord)tr.GetObject(btrId, OpenMode.ForRead);
                    if (btr.IsLayout) continue;

                    foreach (ObjectId entId in btr)
                    {
                        Entity ent = tr.GetObject(entId, OpenMode.ForRead) as Entity;
                        if (ent != null && ent.Layer.Equals(targetLayer, StringComparison.OrdinalIgnoreCase))
                        {
                            ent.UpgradeOpen();
                            ent.Layer = "0";
                            changedCount++;
                        }
                    }
                }

                tr.Commit();
                ed.WriteMessage(string.Format("\nУспешно очищено объектов: {0}. Теперь вы можете удалить слой штатными средствами.", changedCount));
            }

            ed.Regen();
        }
    }
}
