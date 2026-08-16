using System;
using System.Windows;
using Application = VMS.TPS.Common.Model.API.Application; // apelido: evita ambiguidade com System.Windows.Application
using Patient = VMS.TPS.Common.Model.API.Patient;
using VMS.TPS.Common.Model.API;
using VMS.TPS.Common.Model.Types;

namespace wpftest
{
    public partial class MainWindow : Window
    {
        // Objeto principal do ESAPI — so pode existir 1 por execucao do programa.
        private Application app;

        // Campo da classe (nao variavel local!) — e por isso que fica "vivo"
        // e acessivel em QUALQUER metodo da janela, inclusive no clique de
        // outro botao. Comeca null ate algum paciente ser aberto com sucesso.
        private Patient pacienteAtual;

        public MainWindow()
        {
            InitializeComponent();

            try
            {
                app = Application.CreateApplication();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erro ao conectar no ESAPI: " + ex.Message);
                Close();
                return;
            }

            // Precisa liberar (Dispose) a Application quando a janela fecha.
            Closing += (s, e) => app.Dispose();
        }

        private void btnAbrirPaciente_Click(object sender, RoutedEventArgs e)
        {
            string id = txtBuscaId.Text.Trim();

            if (string.IsNullOrEmpty(id))
            {
                lblStatus.Text = "Digite um ID de paciente.";
                return;
            }

            try
            {
                // So pode ter 1 paciente aberto por vez — fecha o anterior, se houver.
                try { app.ClosePatient(); } catch { /* nao tinha paciente aberto ainda, tudo bem */ }

                // Atribui ao CAMPO da classe (pacienteAtual), nao a uma variavel
                // local — assim outros metodos/botoes conseguem usar depois.
                pacienteAtual = app.OpenPatientById(id);

                if (pacienteAtual != null)
                {
                    lblStatus.Text = "Paciente " + pacienteAtual.Id + " aberto com sucesso.";
                    btnIniciarAutomacao.IsEnabled = true; // so libera automacao com paciente aberto
                }
                else
                {
                    lblStatus.Text = "Paciente não encontrado.";
                    btnIniciarAutomacao.IsEnabled = false;
                }
            }
            catch (Exception ex)
            {
                lblStatus.Text = "Erro ao abrir paciente: " + ex.Message;
                btnIniciarAutomacao.IsEnabled = false;
            }
        }

        private void btnIniciarAutomacao_Click(object sender, RoutedEventArgs e)
        {
            // Aqui usamos o CAMPO pacienteAtual — o mesmo objeto que foi
            // aberto no clique do botao anterior. Essa checagem e uma
            // segunda garantia (o botao ja comeca desabilitado sem paciente,
            // mas e uma boa pratica nao confiar so nisso).
            if (pacienteAtual == null)
            {
                lblStatus.Text = "Abra um paciente antes de iniciar a automação.";
                return;
            }

            lblStatus.Text = "Iniciando automação para o paciente " + pacienteAtual.Id + "...";

            // A partir daqui e so usar "pacienteAtual" normalmente, por
            // exemplo: pacienteAtual.Courses, pacienteAtual.StructureSets etc.
        }
    }
}