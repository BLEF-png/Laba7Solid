using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System.Runtime.InteropServices;

namespace Lab5_Kaluzhny
{
    public class Remover
    {
        private static Feature _feature;

        /// <summary>
        /// Полная зачистка "хвоста" модели до 5 первых фич.
        /// Используется один раз при новом запуске расчёта.
        /// </summary>
        public static void RemoveFeature(ModelDoc2 modelDoc2)
        {
            if (modelDoc2 == null) return;

            try
            {
                while (true)
                {
                    int count;
                    try
                    {
                        count = modelDoc2.GetFeatureCount();
                    }
                    catch (COMException)
                    {
                        // Солид отдал RPC_E_DISCONNECTED – просто выходим
                        return;
                    }

                    if (count <= 5)
                        return;

                    // если StepRemove не смог ничего удалить – выходим, чтобы не зациклиться
                    if (!StepRemoveInternal(modelDoc2))
                        return;
                }
            }
            catch (COMException)
            {
                // на всякий пожарный – больше ничего не делаем, чтобы не ронять программу
            }
        }

        /// <summary>
        /// Внешний метод: удалить одну "последнюю" итерацию (вырез + эскиз).
        /// </summary>
        public static void StepRemove(ModelDoc2 modelDoc2)
        {
            StepRemoveInternal(modelDoc2);
        }

        /// <summary>
        /// Внутренняя логика удаления, вернёт false, если ничего удалить не удалось
        /// (например, из-за COMException или отсутствия фич).
        /// </summary>
        private static bool StepRemoveInternal(ModelDoc2 modelDoc2)
        {
            if (modelDoc2 == null) return false;

            try
            {
                // --- 1) последний feature ---
                _feature = modelDoc2.FeatureByPositionReverse(0) as Feature;
            }
            catch (COMException)
            {
                return false;
            }

            if (_feature == null) return false;

            string name;
            string type2;

            try
            {
                name = _feature.Name ?? "";
                type2 = _feature.GetTypeName2() ?? _feature.GetTypeName();
            }
            catch (COMException)
            {
                return false;
            }

            // не трогаем служебную геометрию
            if (name == "RegionTopPlane" || name == "RegionTopPtsSketch" ||
                type2 == "RefPlane" || type2 == "3DProfileFeature" || type2 == "3DSketch")
                return false;

            // если это не эскиз – это вырез / бобышка → удалить
            if (type2 != "ProfileFeature")
            {
                try
                {
                    modelDoc2.ClearSelection2(true);
                    ((Entity)_feature).Select2(false, 0);
                    modelDoc2.EditDelete();
                }
                catch (COMException)
                {
                    return false;
                }
            }

            // --- 2) теперь на хвосте, как правило, эскиз этого выреза ---
            try
            {
                _feature = modelDoc2.FeatureByPositionReverse(0) as Feature;
            }
            catch (COMException)
            {
                return false;
            }

            if (_feature == null) return false;

            try
            {
                name = _feature.Name ?? "";
                type2 = _feature.GetTypeName2() ?? _feature.GetTypeName();
            }
            catch (COMException)
            {
                return false;
            }

            if (type2 == "ProfileFeature" && name != "RegionTopPtsSketch")
            {
                try
                {
                    modelDoc2.ClearSelection2(true);
                    ((Entity)_feature).Select2(false, 0);
                    modelDoc2.EditDelete();
                }
                catch (COMException)
                {
                    return false;
                }
            }

            return true;
        }
    }
}
