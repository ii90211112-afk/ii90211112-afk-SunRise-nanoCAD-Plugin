using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using HostMgd.ApplicationServices; // Для HostApplication
using System.Management;
using Teigha.Runtime;

// 1. Обязательный атрибут, который указывает nanoCAD на ваш стартовый класс
[assembly: ExtensionApplication(typeof(SunRise.SunRise))]

namespace SunRise
{
    using System.Reflection;
    using System.Runtime.InteropServices;
    using System.Text.RegularExpressions;
    using HostMgd.EditorInput;
    using Teigha.DatabaseServices;
    using Teigha.Runtime;

    public class SunRise: IExtensionApplication
    {
        public void Initialize()
        {
            var ed = Application.DocumentManager.MdiActiveDocument.Editor;
            ed.WriteMessage("Расширение SunRise v:0.0.0.1 успешно загружено");
            ed.WriteMessage("Введите команду SunRiseInfo, для получения подробностей");
        }

        public void Terminate()
        {
            // Здесь можно освобождать ресурсы при закрытии nanoCAD
        }

        [CommandMethod("SunRiseInfo")]
        public void Info()
        {
            var ed = Application.DocumentManager.MdiActiveDocument.Editor;
            ed.WriteMessage("Расширение SunRise версия 0.0.0.1"); 
            ed.WriteMessage("Доступные команды: ");
            ed.WriteMessage("CheckDisk - проверяет СМАРТ-статуса диска, может не работать без админских прав");
            ed.WriteMessage("PaintItBlack - красит все объекты в черный цвет, даже внутри блоков может быть необходимо для печати. Примечание: чтобы перекрасить МТекст-ы, нужно их разбить");
            ed.WriteMessage("ConvertToMText - преобразует текст(или тексты) в МТекст. При указании более одного текста - объединяет их в один МТекст");
            ed.WriteMessage("FixTrueColors - преобразует цвета из х,х,х в цвета NanoCad - особо часто требуется для печати");
            ed.WriteMessage("FlattenDrawing - ");

        }

        /// <summary>
        /// Команда для проверки СМАРТ-статуса диска, может не работать без админских прав
        /// </summary>
        [CommandMethod ("CheckDisk")]
        public static void CheckStatus()
        {
            var ed = Application.DocumentManager.MdiActiveDocument.Editor;
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
            var ed = Application.DocumentManager.MdiActiveDocument.Editor;
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

            Document doc = Application.DocumentManager.MdiActiveDocument;
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
            var doc = Application.DocumentManager.MdiActiveDocument;
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
            var doc = Application.DocumentManager.MdiActiveDocument;
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




    }
}
