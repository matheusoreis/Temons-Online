import 'package:servidor/interfaces/receiver_message_interface.dart';
import 'package:servidor/models/connection_model.dart';
import 'package:servidor/network/data/senders/messages/sender_ping_message.dart';
import 'package:servidor/network/server_buffer.dart';

///
class ReceiverPingMessage implements ReceiverMessageInterface {
  final ServerBuffer _buffer = ServerBuffer();
  final SenderPingMessage _sender = SenderPingMessage();

  @override
  void receiver({
    required ConnectionModel client,
    required List<int> data,
  }) {
    _buffer.writeBytes(data);

    _sender(client: client);
  }
}
