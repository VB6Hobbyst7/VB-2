unit U_QC_PRT;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, QuickRpt, QRCtrls, ExtCtrls;

type
  TF_QC_PRT = class(TForm)
    QuickRep1: TQuickRep;
    QRBand1: TQRBand;
    qrlTitle: TQRLabel;
    qrlTestGG: TQRLabel;
    qrlTestDt: TQRLabel;
    QRBand2: TQRBand;
    qrlGrp: TQRImage;
    QRShape1: TQRShape;
    QRLabel21: TQRLabel;
    QRLabel23: TQRLabel;
    QRLabel24: TQRLabel;
    T1: TQRLabel;
    L1: TQRLabel;
    R1: TQRLabel;
    QRShape2: TQRShape;
    QRShape3: TQRShape;
    QRShape4: TQRShape;
    QRLabel26: TQRLabel;
    QRLabel27: TQRLabel;
    QRLabel28: TQRLabel;
    QRLabel29: TQRLabel;
    QRLabel31: TQRLabel;
    QRLabel32: TQRLabel;
    T2: TQRLabel;
    L2: TQRLabel;
    R2: TQRLabel;
    T3: TQRLabel;
    L3: TQRLabel;
    R3: TQRLabel;
    T4: TQRLabel;
    L4: TQRLabel;
    R4: TQRLabel;
    T5: TQRLabel;
    L5: TQRLabel;
    R5: TQRLabel;
    T6: TQRLabel;
    L6: TQRLabel;
    R6: TQRLabel;
    T7: TQRLabel;
    L7: TQRLabel;
    R7: TQRLabel;
    T8: TQRLabel;
    L8: TQRLabel;
    R8: TQRLabel;
    T9: TQRLabel;
    L9: TQRLabel;
    R9: TQRLabel;
    T10: TQRLabel;
    L10: TQRLabel;
    R10: TQRLabel;
    T11: TQRLabel;
    L11: TQRLabel;
    R11: TQRLabel;
    T12: TQRLabel;
    L12: TQRLabel;
    R12: TQRLabel;
    T13: TQRLabel;
    L13: TQRLabel;
    R13: TQRLabel;
    T14: TQRLabel;
    L14: TQRLabel;
    R14: TQRLabel;
    T15: TQRLabel;
    L15: TQRLabel;
    R15: TQRLabel;
    T16: TQRLabel;
    L16: TQRLabel;
    R16: TQRLabel;
    T17: TQRLabel;
    L17: TQRLabel;
    R17: TQRLabel;
    T18: TQRLabel;
    L18: TQRLabel;
    R18: TQRLabel;
    T19: TQRLabel;
    L19: TQRLabel;
    R19: TQRLabel;
    R20: TQRLabel;
    L20: TQRLabel;
    T20: TQRLabel;
    T21: TQRLabel;
    L21: TQRLabel;
    R21: TQRLabel;
    R22: TQRLabel;
    L22: TQRLabel;
    T22: TQRLabel;
    T23: TQRLabel;
    L23: TQRLabel;
    R23: TQRLabel;
    R24: TQRLabel;
    L24: TQRLabel;
    T24: TQRLabel;
    T25: TQRLabel;
    L25: TQRLabel;
    R25: TQRLabel;
    R26: TQRLabel;
    L26: TQRLabel;
    T26: TQRLabel;
    T27: TQRLabel;
    L27: TQRLabel;
    R27: TQRLabel;
    R28: TQRLabel;
    L28: TQRLabel;
    T28: TQRLabel;
    T29: TQRLabel;
    L29: TQRLabel;
    R29: TQRLabel;
    R30: TQRLabel;
    L30: TQRLabel;
    T30: TQRLabel;
    T31: TQRLabel;
    L31: TQRLabel;
    R31: TQRLabel;
    R32: TQRLabel;
    L32: TQRLabel;
    T32: TQRLabel;
    T33: TQRLabel;
    L33: TQRLabel;
    R33: TQRLabel;
    R34: TQRLabel;
    L34: TQRLabel;
    T34: TQRLabel;
    T35: TQRLabel;
    L35: TQRLabel;
    R35: TQRLabel;
    R36: TQRLabel;
    L36: TQRLabel;
    T36: TQRLabel;
    T37: TQRLabel;
    L37: TQRLabel;
    R37: TQRLabel;
    R38: TQRLabel;
    L38: TQRLabel;
    T38: TQRLabel;
    T39: TQRLabel;
    L39: TQRLabel;
    R39: TQRLabel;
    R40: TQRLabel;
    L40: TQRLabel;
    T40: TQRLabel;
    T41: TQRLabel;
    L41: TQRLabel;
    R41: TQRLabel;
    R42: TQRLabel;
    L42: TQRLabel;
    T42: TQRLabel;
    T43: TQRLabel;
    L43: TQRLabel;
    R43: TQRLabel;
    R44: TQRLabel;
    L44: TQRLabel;
    T44: TQRLabel;
    T45: TQRLabel;
    L45: TQRLabel;
    R45: TQRLabel;
    R46: TQRLabel;
    L46: TQRLabel;
    T46: TQRLabel;
    T47: TQRLabel;
    L47: TQRLabel;
    R47: TQRLabel;
    R48: TQRLabel;
    L48: TQRLabel;
    T48: TQRLabel;
    T49: TQRLabel;
    L49: TQRLabel;
    R49: TQRLabel;
    R50: TQRLabel;
    L50: TQRLabel;
    T50: TQRLabel;
    T51: TQRLabel;
    L51: TQRLabel;
    R51: TQRLabel;
    R52: TQRLabel;
    L52: TQRLabel;
    T52: TQRLabel;
    T53: TQRLabel;
    L53: TQRLabel;
    R53: TQRLabel;
    R54: TQRLabel;
    L54: TQRLabel;
    T54: TQRLabel;
    qrlENM: TQRLabel;
    qrlUnit: TQRLabel;
    qrlCor: TQRLabel;
    qrlLot: TQRLabel;
    qrlMean: TQRLabel;
    qrlSD: TQRLabel;
    qrlCV: TQRLabel;
    qrl2SD: TQRLabel;
    qrl3SD: TQRLabel;
    QRShape5: TQRShape;
    QRShape6: TQRShape;
    QRLabel22: TQRLabel;
    QRLabel25: TQRLabel;
    QRLabel30: TQRLabel;
    T55: TQRLabel;
    T56: TQRLabel;
    T57: TQRLabel;
    T58: TQRLabel;
    T59: TQRLabel;
    T60: TQRLabel;
    T61: TQRLabel;
    T62: TQRLabel;
    T63: TQRLabel;
    T64: TQRLabel;
    T65: TQRLabel;
    T66: TQRLabel;
    T67: TQRLabel;
    T68: TQRLabel;
    T69: TQRLabel;
    T70: TQRLabel;
    T71: TQRLabel;
    T72: TQRLabel;
    L72: TQRLabel;
    L71: TQRLabel;
    L70: TQRLabel;
    L69: TQRLabel;
    L68: TQRLabel;
    L67: TQRLabel;
    L66: TQRLabel;
    L65: TQRLabel;
    L64: TQRLabel;
    L63: TQRLabel;
    L62: TQRLabel;
    L61: TQRLabel;
    L60: TQRLabel;
    L59: TQRLabel;
    L58: TQRLabel;
    L57: TQRLabel;
    L56: TQRLabel;
    L55: TQRLabel;
    R55: TQRLabel;
    R56: TQRLabel;
    R57: TQRLabel;
    R58: TQRLabel;
    R59: TQRLabel;
    R60: TQRLabel;
    R61: TQRLabel;
    R62: TQRLabel;
    R63: TQRLabel;
    R64: TQRLabel;
    R65: TQRLabel;
    R66: TQRLabel;
    R67: TQRLabel;
    R68: TQRLabel;
    R69: TQRLabel;
    R70: TQRLabel;
    R71: TQRLabel;
    R72: TQRLabel;
    QRLabel87: TQRLabel;
    QRLabel88: TQRLabel;
    QRLabel89: TQRLabel;
    R73: TQRLabel;
    R74: TQRLabel;
    R76: TQRLabel;
    R75: TQRLabel;
    L75: TQRLabel;
    L74: TQRLabel;
    L73: TQRLabel;
    T73: TQRLabel;
    T74: TQRLabel;
    T75: TQRLabel;
    T76: TQRLabel;
    T77: TQRLabel;
    T78: TQRLabel;
    T79: TQRLabel;
    T80: TQRLabel;
    T81: TQRLabel;
    T82: TQRLabel;
    T83: TQRLabel;
    T84: TQRLabel;
    T85: TQRLabel;
    T86: TQRLabel;
    T87: TQRLabel;
    T88: TQRLabel;
    T89: TQRLabel;
    T90: TQRLabel;
    L90: TQRLabel;
    R90: TQRLabel;
    R89: TQRLabel;
    L89: TQRLabel;
    L88: TQRLabel;
    R88: TQRLabel;
    R87: TQRLabel;
    L87: TQRLabel;
    L86: TQRLabel;
    R86: TQRLabel;
    R85: TQRLabel;
    L85: TQRLabel;
    L84: TQRLabel;
    R84: TQRLabel;
    R83: TQRLabel;
    L83: TQRLabel;
    L82: TQRLabel;
    R82: TQRLabel;
    R81: TQRLabel;
    L81: TQRLabel;
    L80: TQRLabel;
    R80: TQRLabel;
    R79: TQRLabel;
    L79: TQRLabel;
    L78: TQRLabel;
    R78: TQRLabel;
    R77: TQRLabel;
    L77: TQRLabel;
    L76: TQRLabel;
    QRShape7: TQRShape;
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  F_QC_PRT: TF_QC_PRT;

implementation

{$R *.dfm}

end.
