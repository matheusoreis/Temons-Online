import 'package:servidor/network/data/packets/server_packets.dart';
import 'package:servidor/network/data/senders/data_sender.dart';
import 'package:servidor/network/server_buffer.dart';
import 'package:servidor/server/server_globals.dart';

///
class SenderHighIndexMessage {
  ///
  void call() {
    _sendHighIndexToAll();
  }

  void _sendHighIndexToAll() {
    final buffer = ServerBuffer();
    final dataSender = DataSender();

    buffer
      ..writeInteger(ServerPackets.sHighIndex.index)
      ..writeInteger(ServerGlobals.playerHighIndex);

    dataSender.sendDataToAll(
      data: buffer.getArray(),
    );

    buffer.flush();
  }
}
