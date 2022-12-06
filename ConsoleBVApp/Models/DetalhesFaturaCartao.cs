namespace ConsoleBVApp.Models
{
    public class DetalhesFaturaCarto
    {
        public string nomeCliente { get; set; }
        public string ultimosQuatroDigitosCartao { get; set; }
        public double valorSubTotalMovimentacaoNacional { get; set; }
        public int valorSubTotalMovimentacaoInternacional { get; set; }
        public List<object> pagamentos { get; set; }
        public List<MovimentacoesNacionai> movimentacoesNacionais { get; set; }
        public List<object> movimentacoesInternacionais { get; set; }
        public bool titular { get; set; }
        public bool ativo { get; set; }
    }

    public class LimitesCreditoFaturaCartao
    {
        public int valorLimiteCreditoRotativoNacional { get; set; }
        public int valorLimiteCreditoRotativoInternacional { get; set; }
        public int valorLimiteSaqueNacional { get; set; }
        public int valorLimiteSaqueInternacional { get; set; }
    }

    public class ListaCamposDinamico
    {
        public int numeroCodigo { get; set; }
        public string descricao { get; set; }
        public double debitoProximoPeriodo { get; set; }
        public double debitoFuturo { get; set; }
    }

    public class MovimentacoesNacionai
    {
        public string nomeCliente { get; set; }
        public DateTime dataMovimentacao { get; set; }
        public string nomeEstabelecimento { get; set; }
        public int numeroParcelaMovimentacao { get; set; }
        public int quantidadeTotalParcelas { get; set; }
        public string nomeCidadeEstabelecimento { get; set; }
        public decimal valorAbsolutoMovimentacaoReal { get; set; }
        public int valorAbsolutoMovimentacaoInternacionalDolar { get; set; }
        public string sinal { get; set; }
        public string moedaMovimentacao { get; set; }
        public int taxaFinanceira { get; set; }
        public int codigo { get; set; }
        public int taxaConversao { get; set; }
        public string moedaLocal { get; set; }
        public string moedaLocalSimbolo { get; set; }
        public int montanteOriginal { get; set; }
        public string moedaOriginal { get; set; }
        public double montanteMoedaOriginal { get; set; }
        public string apelidoMoedaOriginal { get; set; }
        public string simboloMoedaOriginal { get; set; }
        public string moedaReferencia { get; set; }
        public string simboloMoedaReferencia { get; set; }
        public string rubro { get; set; }
        public string estadoTransacao { get; set; }
        public string descTipoTrans { get; set; }
    }

    public class OpcoesParcelamento
    {
        public double taxaMensal { get; set; }
        public double taxaAnual { get; set; }
        public double taxaCETAnual { get; set; }
        public double valorTotalParcelamento { get; set; }
        public int quantidadeParcelas { get; set; }
        public double valorParcela { get; set; }
    }

    public class Fatura
    {
        public int elegibilidadeCompulsorio { get; set; }
        public double encargoRotativosMensal { get; set; }
        public double iofMes { get; set; }
        public double iofAdicional { get; set; }
        public double cetMensal { get; set; }
        public double cetAnual { get; set; }
        public double acrescimoJuros { get; set; }
        public double taxaAtual { get; set; }
        public double taxaPromocional { get; set; }
        public string dataVencimentoFatura { get; set; }
        public string dataHoraGeracaoFatura { get; set; }
        public decimal valorTotalFatura { get; set; }
        public double valorTotalFaturaAnterior { get; set; }
        public double valorPagamentoFaturaAnterior { get; set; }
        public int valorSaldoAnterior { get; set; }
        public int valorEncargos { get; set; }
        public double valorAjustes { get; set; }
        public double valorTotalMovimentacaoNacional { get; set; }
        public decimal valorTotalNacional { get; set; }
        public int valorTotalMovimentacaoInternacionalReal { get; set; }
        public int valorTotalMovimentacaoInternacionalDolar { get; set; }
        public string dataCotacaoDolarFatura { get; set; }
        public double valorCotacaoDolarFatura { get; set; }
        public double valorTotalComprasParceladasFutura { get; set; }
        public string dataVencimentoProximaFatura { get; set; }
        public decimal valorPagamentoMinimo { get; set; }
        public int valorVencimentoAlternativo { get; set; }
        public int quantidadeParcelasPagamento { get; set; }
        public double valorParcelaPagamento { get; set; }
        public List<OpcoesParcelamento> opcoesParcelamento { get; set; }
        public LimitesCreditoFaturaCartao limitesCreditoFaturaCartao { get; set; }
        public TaxasFaturaCartao taxasFaturaCartao { get; set; }
        public List<DetalhesFaturaCarto> detalhesFaturaCartoes { get; set; }
        public string dataInicioMovimentacoes { get; set; }
        public string dataTerminoMovimentacoes { get; set; }
        public string dataProximoCorte { get; set; }
        public List<ListaCamposDinamico> listaCamposDinamicos { get; set; }
    }

    public class TaxasFaturaCartao
    {
        public int taxaMediaPonderada { get; set; }
        public double taxaSaldoFinanciadoPeriodoMensal { get; set; }
        public double taxaSaldoFinanciadoPeriodoAnual { get; set; }
        public double taxaSaqueNacionalPeriodoMensal { get; set; }
        public double taxaSaqueNacionalPeriodoAnual { get; set; }
        public double taxaSaldoFinanciadoProximoPeriodoMensal { get; set; }
        public double taxaSaldoFinanciadoProximoPeriodoAnual { get; set; }
        public double taxaSaqueNacionalProximoPeriodoMensal { get; set; }
        public double taxaSaqueNacionalProximoPeriodoAnual { get; set; }
        public double taxaIofFinanciamentoValor { get; set; }
        public double taxaIofFinanciamentoMensal { get; set; }
        public double taxaIofTransacaoInternacional { get; set; }
        public double taxaCetSaldoFinanciadoMensal { get; set; }
        public double taxaCetSaldoFinanciadoAnual { get; set; }
        public double taxaCetSaqueNacionalMensal { get; set; }
        public double taxaCetSaqueNacionalAnual { get; set; }
        public double valorTarifaSaque { get; set; }
        public int taxaCetParcelamentoFaturaMensal { get; set; }
        public int taxaCetParcelamentoFaturaAnual { get; set; }
        public double taxaCetPagamentoContaMensal { get; set; }
        public double taxaCetPagamentoContaAnual { get; set; }
        public double valorTarifaPagamentoConta { get; set; }
        public double taxaCetParcelamentoEmissorMensal { get; set; }
        public double taxaCetParcelamentoEmissorAnual { get; set; }
        public int taxaParcelamentoPeriodoMensal { get; set; }
        public int taxaParcelamentoPeriodoAnual { get; set; }
        public int taxaParcelamentoSaldoTotalPeriodoMensal { get; set; }
        public int taxaParcelamentoSaldoTotalPeriodoAnual { get; set; }
        public int taxaParcelamentoProximoPeriodoMensal { get; set; }
        public int taxaParcelamentoProximoPeriodoAnual { get; set; }
        public int taxaParcelamentoSaldoTotalProximoPeriodoMensal { get; set; }
        public int taxaParcelamentoSaldoTotalProximoPeriodoAnual { get; set; }
        public int taxaCetParcelamentoSaldoTotalMensal { get; set; }
        public int taxaCetParcelamentoSaldoTotalAnual { get; set; }
        public decimal taxaClientefaturaCartaoPrestacao { get; set; }
        public int taxaClienteFundos { get; set; }
        public double taxaClienteEspecialfaturaCartaoPrestacao { get; set; }
        public double taxaClienteSaquePrestacao { get; set; }
        public double taxafaturaCartaoPrestacao { get; set; }
        public double taxaFundos { get; set; }
        public double proximaAnualCETPrestacao { get; set; }
        public double proximaAnualCETEspecialPrestacao { get; set; }
        public double proximaAnualCETSaquePrestacao { get; set; }
        public decimal proximaTaxaClientefaturaCartaoPrestacao { get; set; }
        public int proximaTaxaClienteFundos { get; set; }
        public double proximaTaxaClienteEspecialfaturaCartaoPrestacao { get; set; }
        public double proximaTaxaClienteSaquePrestacao { get; set; }
        public double proximaTaxafaturaCartaoPrestacao { get; set; }
        public double proximaTaxaFundos { get; set; }
        public double proximaMensalCETPrestacao { get; set; }
        public double proximaMensalCETEspecialPrestacao { get; set; }
        public double proximaMensalCETSaquePrestacao { get; set; }
        public double proximaTaxaEspecialfaturaCartaoPrestacao { get; set; }
        public double proximaTaxaSaquePrestacao { get; set; }
        public double taxaEspecialfaturaCartaoPrestacao { get; set; }
        public double taxaSaquePrestacao { get; set; }
    }
}
