using Stage2WPF.Core;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Stage2WPF.MVVM.ViewModel
{
    class MainViewModel : ObservableObject
    {
        public RelayCommand HomeViewCommand { get; set; }
        public RelayCommand ExcelViewCommand { get; set; }

        private object _currentView;
        public HomeViewModel HomeVM { get; set; }
        public ExcelViewModel ExcelVM { get; set; }

        public object CurrentView
        {
            get { return _currentView; }
            set
            {
                _currentView = value;
                OnPropertyChanged();
            }
        }

        public MainViewModel()
        {
            HomeVM = new HomeViewModel();
            ExcelVM = new ExcelViewModel();
            CurrentView = HomeVM;

            HomeViewCommand = new RelayCommand(o =>
            {
                CurrentView = HomeVM;
            });

            ExcelViewCommand = new RelayCommand(o =>
            {
                CurrentView = ExcelVM;
            });
        }
    }
}
