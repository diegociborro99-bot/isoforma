// Vitest config: Node environment, pattern tests/**/*.test.js
module.exports = {
  test: {
    environment: 'node',
    include: ['tests/**/*.test.js'],
    reporters: ['verbose'],
    testTimeout: 10000
  }
};
