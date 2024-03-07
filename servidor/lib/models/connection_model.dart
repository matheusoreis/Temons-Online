import 'dart:io';

/// modelos de dados de conexão do cliente.
///
/// Este modelo representa uma instância de cliente de rede.
class ConnectionModel {
  /// Cria uma nova instância de `ClientConnectionModel` com os dados do cliente.
  ConnectionModel({required this.id, required this.socket});

  /// Identificador do cliente.
  final int id;

  /// Socket do cliente.
  final Socket socket;
}
