import 'package:servidor/models/connection_model.dart';
import 'package:servidor/network/server_buffer.dart';
import 'package:servidor/server/server_constants.dart';
import 'package:servidor/server/server_memory.dart';

///
class PlayerModel extends ConnectionModel {
  ///
  PlayerModel({
    required super.id,
    required super.socket,
    this.inGame,
  });

  ///
  bool? inGame = false;

  ///
  ServerBuffer buffer = ServerBuffer();

  /// Verifica se um determinado índice de slot está atualmente conectado.
  ///
  /// Este método retorna verdadeiro se o slot especificado estiver ocupado por um cliente conectado,
  /// caso contrário, retorna falso.
  ///
  /// Retorna verdadeiro se o slot estiver ocupado, falso caso contrário.
  bool isConnected() {
    if (id < 0 || id >= ServerConfig.maxPlayers) {
      return false;
    }

    return !ServerMemory().clientConnections.isSlotEmpty(id);
  }

  ///
  bool isPlaying() {
    if (id < 0 || id >= ServerConfig.maxPlayers) {
      return false;
    }

    return inGame ?? false;
  }
}
