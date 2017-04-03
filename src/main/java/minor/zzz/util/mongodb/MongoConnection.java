package minor.zzz.util.mongodb;

import com.mongodb.MongoClient;
import com.mongodb.ServerAddress;

import java.util.concurrent.ConcurrentHashMap;

/**
 * MongoClient自身是实现了连接池的
 * Created by admin on 2017/3/31.
 */
public class MongoConnection {

    // 这里强调下，ServerAddress里本身重写了hashCode方法
    private static final ConcurrentHashMap<ServerAddress, MongoClient> connMap = new ConcurrentHashMap<>();

    public static final MongoClient getConnection(String host, int port) {

        ServerAddress address = new ServerAddress(host, port);

        MongoClient client = connMap.get(address);

        if (client == null) {
            client = MongoConnectionFactory.getConnection(address);
            connMap.putIfAbsent(address, client);
        }

        return connMap.get(address);
    }

    private static class MongoConnectionFactory {

        public static final MongoClient getConnection(String host, int port) {

            return new MongoClient(host, port);
        }

        public static final MongoClient getConnection(ServerAddress address) {
            return new MongoClient(address);
        }
    }

}
