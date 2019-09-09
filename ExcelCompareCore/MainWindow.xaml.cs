using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace ExcelCompareCore
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }
    }


    public abstract class ModifierBase
    {
        public abstract void Apply(DependencyObject target);
    }

    public class Modifier : ModifierBase
    {
        public DependencyProperty Property { get; set; }
        public object Value { get; set; }
        public string TemplatePartName { get; set; }
        public override void Apply(DependencyObject target)
        {
            if (target is FrameworkElement element &&
                target.GetValue(Control.TemplateProperty) is ControlTemplate template &&
                template.FindName(TemplatePartName, element) is DependencyObject templatePart)
            {
                templatePart.SetCurrentValue(Property, Value);
            }
        }
    }

    public class ModifierCollection : Collection<Modifier>
    { }

    public class TreeHelpers
    {
        public static readonly DependencyProperty ModifiersProperty = DependencyProperty.RegisterAttached(
            "Modifiers", typeof(ModifierCollection), typeof(TreeHelpers), new PropertyMetadata(default(ModifierCollection), PropertyChangedCallback));

        private static void PropertyChangedCallback(DependencyObject dependencyObject, DependencyPropertyChangedEventArgs e)
        {
            if (dependencyObject is FrameworkElement element && !element.IsLoaded)
            {
                element.Loaded += ElementOnLoaded;
            }
            else
            {
                ApplyModifiers(e.NewValue as IEnumerable<ModifierBase>);
            }

            void ApplyModifiers(IEnumerable<ModifierBase> modifiers)
            {
                foreach (var modifier in modifiers ?? Enumerable.Empty<ModifierBase>())
                {
                    modifier.Apply(dependencyObject);
                }
            }

            void ElementOnLoaded(object sender, RoutedEventArgs routedEventArgs)
            {
                ((FrameworkElement)sender).Loaded -= ElementOnLoaded;
                ApplyModifiers(GetModifiers((FrameworkElement)sender));
            }
        }

        public static void SetModifiers(DependencyObject element, ModifierCollection value)
        {
            element.SetValue(ModifiersProperty, value);
        }

        public static ModifierCollection GetModifiers(DependencyObject element)
        {
            return (ModifierCollection)element.GetValue(ModifiersProperty);
        }
    }
}
