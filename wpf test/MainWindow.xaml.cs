using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls; // TextChangedEventArgs / SelectionChangedEventArgs (eventos da caixa de busca e da lista de sugestoes)
using Application = VMS.TPS.Common.Model.API.Application;   // apelido: evita ambiguidade com System.Windows.Application
using Patient = VMS.TPS.Common.Model.API.Patient;
using VMS.TPS.Common.Model.API;
using VMS.TPS.Common.Model.Types;

[assembly:ESAPIScript(IsWriteable = true)] // permite escrever no banco do Aria

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

        // Lista de TODOS os pacientes, carregada 1 UNICA VEZ quando o programa
        // abre. Diferente do banco SQL do AriaQ (que fica na rede - por isso
        // la usamos timer de espera + busca assincrona), essa lista fica
        // pronta na memoria do programa. Filtrar uma lista que ja esta na
        // memoria e muito rapido, entao da pra fazer isso a cada tecla
        // digitada sem se preocupar com performance nem usar threads.
        private List<PatientSummary> todosOsPacientes;

        public MainWindow()
        {
            InitializeComponent();

            try
            {
                app = Application.CreateApplication();
                todosOsPacientes = app.PatientSummaries.ToList();
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

        // Dispara a CADA tecla digitada na caixa de ID. So filtra a lista que
        // ja esta na memoria (todosOsPacientes) - por isso pode rodar direto
        // aqui, sem timer de espera nem thread separada.
        private void txtBuscaId_TextChanged(object sender, TextChangedEventArgs e)
        {
            string termo = txtBuscaId.Text.Trim();
            lstSugestoes.Items.Clear();

            if (termo.Length == 0 || todosOsPacientes == null) return;

            var resultados = todosOsPacientes
                .Where(p => p.Id != null && p.Id.StartsWith(termo, StringComparison.OrdinalIgnoreCase))
                .OrderBy(p => p.Id)
                .Take(15) // limita a lista pra nao ficar gigante na tela
                .ToList();

            foreach (var p in resultados)
                lstSugestoes.Items.Add(p.Id + " - " + p.LastName + ", " + p.FirstName);
        }

        // Ao clicar numa sugestao, so preenche o campo de ID - o fluxo de
        // abrir o paciente continua sendo o mesmo botao "Abrir paciente" de
        // sempre, isso so ajuda a achar o ID certo mais rapido.
        private void lstSugestoes_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (lstSugestoes.SelectedItem == null) return;

            string selecionado = lstSugestoes.SelectedItem.ToString();
            txtBuscaId.Text = selecionado.Split('-')[0].Trim(); // pega so o ID, antes do " - "
            lstSugestoes.Items.Clear();
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

        // ---------------------------------------------------------------
        // AUTOMAÇÃO DE PLANO (adaptado do script de exemplo da Varian pro
        // Halcyon). A lógica clínica (prescrição, objetivos de otimização,
        // campos, modelos de cálculo) foi mantida IGUAL ao script original
        // — não mudei nenhum número/parâmetro clínico. O que mudou foi só
        // a "cola" em volta, pra encaixar no nosso projeto:
        //
        //   1) NÃO abre um paciente novo (removido "OpenPatientById" daqui
        //      dentro) — usa o "pacienteAtual" que já foi aberto pelo botão
        //      "Abrir paciente". Só existe UM paciente aberto por vez no
        //      ESAPI, então reaproveitar evita abrir dois ao mesmo tempo.
        //   2) Troquei a "listBox1" (que não existia no nosso XAML) por uma
        //      caixa de texto simples (txtResultados).
        //   3) Unifiquei todo o log em Trace.WriteLine só (o script original
        //      misturava Console.WriteLine e Trace.WriteLine — os dois
        //      aparecem na mesma janela de console, então usar só um é mais
        //      simples de entender).
        //   4) Coloquei tudo dentro de um try/catch/finally, pra garantir
        //      que o paciente seja fechado e o botão travado de novo mesmo
        //      se der erro no meio do caminho.
        //
        // *** ATENÇÃO — ISSO ESCREVE DE VERDADE NO BANCO DO ARIA ***
        // Cria curso, plano, roda otimização e calcula dose. Só rode isso
        // numa base de TESTE (T-Box), nunca em produção.
        // ---------------------------------------------------------------
        private void btnIniciarAutomacao_Click(object sender, RoutedEventArgs e)
        {
            if (pacienteAtual == null)
            {
                lblStatus.Text = "Abra um paciente antes de iniciar a automação.";
                return;
            }

            // Log de acompanhamento: aparece na janela "Output" do Visual
            // Studio enquanto você roda com F5 (Debug > janelas > Output,
            // ou já vem aberta por padrão). Tiramos o console preto
            // (AllocConsole) porque ele tinha um problema chato de
            // compatibilidade com o debugger do Visual Studio — o Trace já
            // resolve sem precisar de janela extra nenhuma.
            Trace.WriteLine("Iniciando a automacao.\n");

            try
            {
                // Obrigatório chamar ANTES de qualquer alteração no paciente.
                pacienteAtual.BeginModifications();

                // ---- Curso ----
                // Procura um curso chamado "HalcyonCourse". Se já existir
                // (de uma execução anterior), apaga e recria do zero — assim
                // não fica acumulando curso duplicado a cada teste.
                Course halcyonCourse = pacienteAtual.Courses.FirstOrDefault(c => c.Id.Equals("HalcyonCourse"));
                if (halcyonCourse != null)
                {
                    pacienteAtual.RemoveCourse(halcyonCourse);
                    Trace.WriteLine("Curso HalcyonCourse existente foi removido.\n");
                }
                halcyonCourse = pacienteAtual.AddCourse();
                halcyonCourse.Id = "HalcyonCourse";
                Trace.WriteLine("Curso criado: " + halcyonCourse.Id + "\n");

                // ---- Structure Set ----
                // DIFERENTE DA VERSÃO ANTERIOR: agora o HalcyonStructSet é
                // sempre APAGADO e RECRIADO do zero, igual ao Curso logo
                // acima. Antes, se já existisse um HalcyonStructSet de uma
                // execução passada, ele era reaproveitado — e aí a mesa
                // (couch), que já tinha sido adicionada da vez anterior,
                // dava erro "Support structures already exist" ao tentar
                // adicionar de novo. Recriando do zero, a automação pode
                // ser rodada várias vezes seguidas sem esse problema.
                StructureSet currentStructureSet = pacienteAtual.StructureSets.FirstOrDefault(s => s.Id == "HalcyonStructSet");
                if (currentStructureSet != null)
                {
                    currentStructureSet.Delete();
                    Trace.WriteLine("HalcyonStructSet existente foi apagado.\n");
                }

                // Copia de um StructureSet AINDA NÃO APROVADO (não dá pra
                // copiar um já aprovado) e renomeia pra "HalcyonStructSet".
                var origem = pacienteAtual.StructureSets
                    .Where(s => s.Structures != null && s.Structures.Any())
                    .FirstOrDefault(s => s.Structures.FirstOrDefault()?.IsApproved == false);

                if (origem != null)
                {
                    currentStructureSet = origem.Copy();
                    currentStructureSet.Id = "HalcyonStructSet";
                    Trace.WriteLine("HalcyonStructSet criado a partir de um StructureSet nao aprovado.\n");
                }
                else
                {
                    Trace.WriteLine("Nenhum StructureSet nao aprovado encontrado. Crie um e tente de novo.\n");
                    app.SaveModifications();
                    return;
                }
                app.SaveModifications();

                // ---- PTV ----
                // Procura uma estrutura com "PTV" no nome e renomeia pra "HalcyonPTV".
                Structure structurePTV = currentStructureSet.Structures.FirstOrDefault(
                    s => s.Id.Contains("PTV") || s.Id.Contains("ptv"));

                if (structurePTV == null)
                {
                    Trace.WriteLine("A estrutura PTV nao existe.\n");
                    return;
                }
                structurePTV.Id = "HalcyonPTV";

                // ---- Mesa (couch) ----
                // Método só existe no ESAPI/Eclipse 16.1 ou mais novo.
                IReadOnlyList<Structure> estruturasAdicionadas;
                bool imagemRedimensionada;
                string mensagemErro;

                currentStructureSet.AddCouchStructures(
                    "RDS_Couch_Top",
                    PatientOrientation.HeadFirstSupine,
                    RailPosition.In,
                    RailPosition.In,
                    null, null, null,
                    out estruturasAdicionadas,
                    out imagemRedimensionada,
                    out mensagemErro);

                if (!string.IsNullOrEmpty(mensagemErro))
                    Trace.WriteLine("Erro ao adicionar a mesa: " + mensagemErro + "\n");

                // ---- Ponto de referência ----
                ReferencePoint halcyonRefPoint = pacienteAtual.ReferencePoints.FirstOrDefault(r => r.Id.Equals("HalcyonRefPoint"));
                if (halcyonRefPoint == null)
                {
                    halcyonRefPoint = pacienteAtual.AddReferencePoint(true, "HalcyonRefPoint");
                    Trace.WriteLine("Ponto de referencia criado.\n");
                }

                // ---- Plano ----
                ExternalPlanSetup halcyonExternalPlan = halcyonCourse.AddExternalPlanSetup(currentStructureSet, structurePTV, halcyonRefPoint);
                halcyonExternalPlan.Id = "HalcyonPlan";
                Trace.WriteLine("Plano criado: " + halcyonExternalPlan.Id + "\n");

                // ---- Prescrição ----
                int numeroFracoes = 5;
                double doseTotalCgy = 3625;
                halcyonExternalPlan.SetPrescription(numeroFracoes, new DoseValue(doseTotalCgy / numeroFracoes, "cGy"), 1);
                Trace.WriteLine("Prescricao definida: " + numeroFracoes + " fracoes, " + doseTotalCgy + " cGy total.\n");

                // ---- Estruturas de risco ----
                // Exemplo pensado pra um caso de próstata — se o paciente de
                // teste não tiver Reto/Bexiga, a automação para aqui de
                // propósito, pra não seguir com dados incompletos.
                Structure structureRectum = currentStructureSet.Structures.FirstOrDefault(s => s.Id.Equals("Rectum"));
                if (structureRectum == null) { Trace.WriteLine("Estrutura Rectum nao encontrada.\n"); return; }

                Structure structureBladder = currentStructureSet.Structures.FirstOrDefault(s => s.Id.Equals("Bladder"));
                if (structureBladder == null) { Trace.WriteLine("Estrutura Bladder nao encontrada.\n"); return; }

                // ---- Objetivos de otimização ----
                // Calculados a partir da dose prescrita (ex.: hot spot =
                // 107% da dose). Mesmos objetivos do script original da Varian.
                double coberturaObjetivo = doseTotalCgy;
                double hotspotObjetivo = doseTotalCgy * 1.07;
                double retoObjetivo1cc = doseTotalCgy * 0.90;
                double retoObjetivo90p = doseTotalCgy * 0.70;
                double bexigaObjetivo1cc = doseTotalCgy * 1.025;
                double bexigaObjetivo90p = doseTotalCgy * 0.60;

                halcyonExternalPlan.OptimizationSetup.AddPointObjective(
                    structurePTV, OptimizationObjectiveOperator.Lower, new DoseValue(coberturaObjetivo, "cGy"), 100, 90);
                halcyonExternalPlan.OptimizationSetup.AddPointObjective(
                    structurePTV, OptimizationObjectiveOperator.Upper, new DoseValue(hotspotObjetivo, "cGy"), 0, 50);

                // Volume relativo equivalente a 1cc, calculado a partir do
                // volume total da estrutura (100% * 1cc / volume total).
                halcyonExternalPlan.OptimizationSetup.AddPointObjective(
                    structureRectum, OptimizationObjectiveOperator.Upper, new DoseValue(retoObjetivo1cc, "cGy"),
                    100 * 1 / structureRectum.Volume, 65);
                halcyonExternalPlan.OptimizationSetup.AddPointObjective(
                    structureRectum, OptimizationObjectiveOperator.Upper, new DoseValue(retoObjetivo90p, "cGy"), 90, 65);

                halcyonExternalPlan.OptimizationSetup.AddPointObjective(
                    structureBladder, OptimizationObjectiveOperator.Upper, new DoseValue(bexigaObjetivo1cc, "cGy"),
                    100 * 1 / structureBladder.Volume, 65);
                halcyonExternalPlan.OptimizationSetup.AddPointObjective(
                    structureBladder, OptimizationObjectiveOperator.Upper, new DoseValue(bexigaObjetivo90p, "cGy"), 90, 65);

                halcyonExternalPlan.OptimizationSetup.AddAutomaticNormalTissueObjective(50);
                Trace.WriteLine("Objetivos de otimizacao adicionados.\n");

                // ---- Campos de tratamento (2 arcos) ----
                ExternalBeamMachineParameters beamMachineParameters =
                    new ExternalBeamMachineParameters("RDSMCH1", "6X", 800, "SRS ARC", "FFF");

                Beam arcBeam1 = halcyonExternalPlan.AddArcBeam(
                    beamMachineParameters, new VRect<double>(-100, -100, 100, 100),
                    330, 181, 179, GantryDirection.Clockwise, 0, structurePTV.CenterPoint);
                Beam arcBeam2 = halcyonExternalPlan.AddArcBeam(
                    beamMachineParameters, new VRect<double>(-100, -100, 100, 100),
                    30, 179, 181, GantryDirection.CounterClockwise, 0, structurePTV.CenterPoint);

                arcBeam1.Id = "HalcyonCW";
                arcBeam2.Id = "HalcyonCCW";
                Trace.WriteLine("Campos adicionados.\n");

                // ---- Imagem de setup ----
                ExternalBeamMachineParameters imagingMachineParameters = new ExternalBeamMachineParameters("RDSMCH1");
                ImagingBeamSetupParameters imagingParameters =
                    new ImagingBeamSetupParameters(ImagingSetup.kVCBCT, 10, 10, 10, 10, 100, 100);
                halcyonExternalPlan.AddImagingSetup(imagingMachineParameters, imagingParameters, structurePTV);
                Trace.WriteLine("Imagem de setup adicionada.\n");

                // ---- Modelos de cálculo (otimização e dose) ----
                halcyonExternalPlan.SetCalculationModel(CalculationType.PhotonVMATOptimization, "PO_1811");
                halcyonExternalPlan.SetCalculationOption("PO_1811", "InhomogeneityCorrection", "On");
                halcyonExternalPlan.SetCalculationOption("PO_1811", "AirCavityCorrection", "On");
                halcyonExternalPlan.SetCalculationOption("PO_1811", "DoseCalculationResolution", "Normal");
                halcyonExternalPlan.SetCalculationOption("PO_1811", "DoseCalculationResolutionForSRSAndHyperarc", "High");
                halcyonExternalPlan.SetCalculationOption("PO_1811", "TargetProjectionMargin", "Normal");
                halcyonExternalPlan.SetCalculationOption("PO_1811", "UseGPU", "Yes");
                halcyonExternalPlan.SetCalculationOption("PO_1811", "AutoFeathering", "On");
                halcyonExternalPlan.SetCalculationOption("PO_1811", "FieldGroupingZThreshold", "3.0");
                halcyonExternalPlan.SetCalculationOption("PO_1811", "ApertureShapeController", "Low");

                halcyonExternalPlan.SetCalculationModel(CalculationType.PhotonVolumeDose, "AXB_1811");
                halcyonExternalPlan.SetCalculationOption("AXB_1811", "HeterogeneityCorrection", "ON");
                halcyonExternalPlan.SetCalculationOption("AXB_1811", "CalculationGridSizeInCM", "0.2");
                halcyonExternalPlan.SetCalculationOption("AXB_1811", "CalculationGridSizeInCMForSRSAndHyperArc", "0.1");
                halcyonExternalPlan.SetCalculationOption("AXB_1811", "FieldNormalizationType", "100% to isocenter");
                app.SaveModifications();

                // ---- Otimização e cálculo de dose (as duas linhas mais demoradas) ----
                Trace.WriteLine("Otimizando (pode demorar alguns minutos)...\n");
                OptimizerResult resultadoOtimizacao = halcyonExternalPlan.OptimizeVMAT(new OptimizationOptionsVMAT(1, string.Empty));

                // NOVO: antes a automação seguia direto pro cálculo de dose mesmo
                // que a otimização tivesse falhado - por isso aparecia o erro
                // "MLC required in Field ...": sem uma otimização bem-sucedida,
                // os campos não têm abertura de MLC válida pra calcular dose em
                // cima. Agora a gente checa "Success" e para aqui se falhar, em
                // vez de continuar com um plano incompleto.
                if (!resultadoOtimizacao.Success)
                {
                    Trace.WriteLine("Otimizacao FALHOU - automacao interrompida antes do calculo de dose.\n");
                    lblStatus.Text = "Otimização falhou. Verifique os parâmetros do plano/campos antes de tentar de novo.";
                    app.SaveModifications();
                    return; // o "finally" ainda roda (fecha o paciente, trava o botão)
                }

                Trace.WriteLine("Otimizacao concluida.\n");
                app.SaveModifications();

                Trace.WriteLine("Calculando a dose...\n");
                halcyonExternalPlan.CalculateDose();
                Trace.WriteLine("Dose calculada.\n");
                app.SaveModifications();

                // ---- Validação do plano ----
                List<PlanValidationResultEsapiDetail> motivos;
                if (!halcyonExternalPlan.IsValidForPlanApproval(out motivos))
                {
                    string mensagem = string.Join("\n", motivos.Select(m => m.MessageForUser));
                    MessageBox.Show("O plano NÃO é válido para aprovação:\n" + mensagem);
                }
                else
                {
                    MessageBox.Show("O plano é válido para aprovação.");
                }

                // ---- Normalização ----
                DoseValue normalizacao = halcyonExternalPlan.GetDoseAtVolume(
                    structurePTV, 95, VolumePresentation.Relative, DoseValuePresentation.Relative);
                halcyonExternalPlan.PlanNormalizationValue = normalizacao.Dose;
                Trace.WriteLine("Normalizacao definida: " + normalizacao.Dose + "%\n");

                // ---- Resumo dos resultados dosimétricos ----
                // Isolado num método próprio (MontarResumoDosimetrico) só
                // pra não deixar esse método gigante ainda maior.
                txtResultados.Text = MontarResumoDosimetrico(
                    halcyonExternalPlan, structurePTV, structureRectum, structureBladder,
                    coberturaObjetivo, hotspotObjetivo,
                    retoObjetivo1cc, retoObjetivo90p,
                    bexigaObjetivo1cc, bexigaObjetivo90p);

                app.SaveModifications();
                Trace.WriteLine("Automacao concluida.\n");
                lblStatus.Text = "Automação concluída para o paciente " + pacienteAtual.Id + ".";
            }
            catch (Exception ex)
            {
                Trace.WriteLine("Erro durante a automacao: " + ex.Message + "\n");
                lblStatus.Text = "Erro durante a automação: " + ex.Message;
            }
            finally
            {
                // Fecha o paciente e trava o botão de novo — depois de
                // fechado, "pacienteAtual" não pode mais ser usado; pra
                // rodar de novo, precisa reabrir pelo primeiro botão.
                app.ClosePatient();
                pacienteAtual = null;
                btnIniciarAutomacao.IsEnabled = false;
            }
        }

        // Monta um texto simples com os valores de dose alcançados x
        // objetivos. Igual à tabela do script original, só que devolvendo
        // uma string em vez de escrever direto numa ListBox.
        private string MontarResumoDosimetrico(ExternalPlanSetup plano, Structure ptv, Structure reto, Structure bexiga,
            double coberturaObjetivo, double hotspotObjetivo,
            double retoObjetivo1cc, double retoObjetivo90p,
            double bexigaObjetivo1cc, double bexigaObjetivo90p)
        {
            DoseValue cobertura = plano.GetDoseAtVolume(ptv, 95, VolumePresentation.Relative, DoseValuePresentation.Absolute);
            DoseValue hotspot = plano.GetDoseAtVolume(ptv, 0.03, VolumePresentation.AbsoluteCm3, DoseValuePresentation.Absolute);
            DoseValue reto1cc = plano.GetDoseAtVolume(reto, 1, VolumePresentation.AbsoluteCm3, DoseValuePresentation.Absolute);
            DoseValue reto90p = plano.GetDoseAtVolume(reto, 90, VolumePresentation.Relative, DoseValuePresentation.Absolute);
            DoseValue bexiga1cc = plano.GetDoseAtVolume(bexiga, 1, VolumePresentation.AbsoluteCm3, DoseValuePresentation.Absolute);
            DoseValue bexiga90p = plano.GetDoseAtVolume(bexiga, 90, VolumePresentation.Relative, DoseValuePresentation.Absolute);

            var sb = new StringBuilder();
            sb.AppendLine("Estrutura\tVolume\tObjetivo\tAlcançado\tPassou?");
            sb.AppendLine(ptv.Id + "\t95%\t" + Math.Round(coberturaObjetivo, 1) + " cGy\t" + Math.Round(cobertura.Dose, 1) + " " + cobertura.Unit + "\t" + (cobertura.Dose > coberturaObjetivo));
            sb.AppendLine(ptv.Id + "\t0.03cc\t" + Math.Round(hotspotObjetivo, 1) + " cGy\t" + Math.Round(hotspot.Dose, 1) + " " + hotspot.Unit + "\t" + (hotspot.Dose < hotspotObjetivo));
            sb.AppendLine(reto.Id + "\t1.00cc\t" + Math.Round(retoObjetivo1cc, 1) + " cGy\t" + Math.Round(reto1cc.Dose, 1) + " " + reto1cc.Unit + "\t" + (reto1cc.Dose < retoObjetivo1cc));
            sb.AppendLine(reto.Id + "\t90%\t" + Math.Round(retoObjetivo90p, 1) + " cGy\t" + Math.Round(reto90p.Dose, 1) + " " + reto90p.Unit + "\t" + (reto90p.Dose < retoObjetivo90p));
            sb.AppendLine(bexiga.Id + "\t1.00cc\t" + Math.Round(bexigaObjetivo1cc, 1) + " cGy\t" + Math.Round(bexiga1cc.Dose, 1) + " " + bexiga1cc.Unit + "\t" + (bexiga1cc.Dose < bexigaObjetivo1cc));
            sb.AppendLine(bexiga.Id + "\t90%\t" + Math.Round(bexigaObjetivo90p, 1) + " cGy\t" + Math.Round(bexiga90p.Dose, 1) + " " + bexiga90p.Unit + "\t" + (bexiga90p.Dose < bexigaObjetivo90p));

            return sb.ToString();
        }
    }
}