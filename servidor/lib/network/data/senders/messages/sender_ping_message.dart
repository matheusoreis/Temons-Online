import 'package:servidor/models/connection_model.dart';
import 'package:servidor/network/data/packets/server_packets.dart';
import 'package:servidor/network/data/senders/data_sender.dart';
import 'package:servidor/network/server_buffer.dart';

///
class SenderPingMessage {
  ///
  void call({
    required ConnectionModel client,
  }) {
    _sendPingTo(client: client);
  }

  void _sendPingTo({
    required ConnectionModel client,
  }) {
    final buffer = ServerBuffer();
    final dataSender = DataSender();

    buffer.writeInteger(ServerPackets.sPing.index);

    dataSender.sendDataTo(
      client: client,
      data: buffer.getArray(),
    );

    buffer.flush();
  }
}
